Attribute VB_Name = "RepGralLiquidacion"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "01/10/2010" 'Martin Ferraro
'Global Const UltimaModificacion = "Version Inicial"

'Global Const Version = "1.01"
'Global Const FechaModificacion = "17/11/2010" 'Martin Ferraro
'Global Const UltimaModificacion = "" 'Banco Buscar de cuenta y no de estructura
                                     'Cuarta parte - TipoCon3 - Buscar los conceptos imprimibles
                                     'Sexta Parte - TipoCon4 - Busca por concepto y acumulador en vez de por tipo de concepto
                                     'Septima parte - No busca mas por tipo de concepto sino por concepto especifico
                                     
'Global Const Version = "1.02"
'Global Const FechaModificacion = "29/11/2010" 'Martin Ferraro
'Global Const UltimaModificacion = "" 'Buscar la ultima fase menor a la fecha del filtro y no la que esta entre las fechas
                                     'Para el total por empresa en los conceptos remunerativos no tomar ni el tipo 2 ni el 3
                                     'Se quito retenciones del total por empresa
                                     'Contribuciones en positivo siempre
                                     'Para el total por empresa en ontribuciones no tomar ni el concepto 11330 ni el concepto 11340
                                     'Provisiones en positivo siempre
                                     
'Global Const Version = "1.03"
'Global Const FechaModificacion = "29/12/2010" 'Martin Ferraro
'Global Const UltimaModificacion = "" 'Los conceptos 11330 y 11340 no sumen a la columna de total CCSS pero si debe sumar a la ultimo columna que informa el total costo empresa
                                     
'Global Const Version = "1.04"
'Global Const FechaModificacion = "24/02/2011" 'Stankunas Cesar
'Global Const UltimaModificacion = "" 'Se modificó la forma de buscar los conceptos para optimizar el reporte

Global Const Version = "1.05"
Global Const FechaModificacion = "10/07/2014" 'Borrelli Facundo
Global Const UltimaModificacion = "" 'Se agregó un "espacio" delante del nro de cuenta bancaria para que en
                                     'la exportacion a Excel, se tome como texto y no como número.
                                     'Se corrige mensaje del log "Buscando CUIL" por "Buscando Datos de la cta. Bancaria"

Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reporte General de Liquidacion.
' Autor      : Martin Ferraro
' Fecha      : 01/10/2010
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
    
    Nombre_Arch = PathFLog & "RHProRepGralLiq-" & NroProcesoBatch & ".log"
    
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
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 274 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call generaRep(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
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


Public Sub generaRep(ByVal BproNro As Long, ByVal Parametros As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera Reporte General de Liquidacion.
' Autor      : Martin Ferraro
' Fecha      : 01/10/2010
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------


Dim CantMov As Integer

'Arreglo que contiene los parametros
Dim arrParam
Dim I As Long

'Parametros desde ASP
Dim ListaProc As String
Dim FecEstr As Date
Dim TituloFiltro As String
Dim BuscarMonto As Boolean


'RecordSet
Dim rs_Empleados As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset
Dim rs_Monto As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset

'Variables
Dim repNro As Long
Dim TECCosto As Long
Dim TEConvenio As Long
Dim TECargaHor As Long
Dim TEPuesto As Long
Dim TECategoria As Long
Dim TETipoPersonal As Long
Dim TESitoLaboral As Long
Dim TEOSocial As Long
Dim TEPlanOSocial As Long
Dim TEBanco As Long
Dim ValCCosto As String
Dim ValConvenio As String
Dim ValCargaHor As String
Dim ValPuesto As String
Dim ValCategoria As String
Dim ValTipoPersonal As String
Dim ValSitoLaboral As String
Dim ValOSocial As String
Dim ValPlanOSocial As String
Dim ValBanco As String
Dim ValSucBanco As String
Dim Columna As Long
Dim TECCostoDesc As String
Dim TEConvenioDesc As String
Dim TECargaHorDesc As String
Dim TEPuestoDesc As String
Dim TECategoriaDesc As String
Dim TETipoPersonalDesc As String
Dim TESitoLaboralDesc As String
Dim TEOSocialDesc As String
Dim TEPlanOSocialDesc As String
Dim TEBancoDesc As String
Dim ArrRemEtiq(1000) As String
Dim ArrRemCod(1000) As Integer
Dim ArrRemTipo(1000) As Integer
Dim IndRem As Long
Dim TotalRem As Double
Dim TotalRemEmpresa As Double
Dim ArrNoRemEtiq(1000) As String
Dim ArrNoRemCod(1000) As Integer
Dim IndNoRem As Long
Dim TotalNoRem As Double
Dim ArrRetEtiq(1000) As String
Dim ArrRetCod(1000) As Integer
Dim IndRet As Long
Dim TotalRet As Double
Dim ArrContrEtiq(1000) As String
Dim ArrContrCod(1000) As Integer
Dim TotalContr As Double
Dim TotalContrEmpresa As Double
Dim IndContr As Long
Dim TipoCon1 As String
Dim TipoCon1Desc As String
Dim TipoCon2 As String
Dim TipoCon2Desc As String
Dim TipoCon3 As String
Dim TipoCon3Desc As String
Dim TipoCon4 As String
Dim TipoCon4Desc As String
Dim TipoCon5 As String
Dim ArrProvEtiq(1000) As String
Dim ArrProvCod(1000) As Integer
Dim IndProv As Long
Dim TotalProv As Double
Dim TotalProvEmpresa As Double
Dim ArrConcAcumEtiq2(1000) As String
Dim ArrConcAcumCod2(1000) As Long
Dim ArrConcAcumTipo2(1000) As Boolean
Dim IndConcAcum2 As Long

Dim ArrConcAcumEtiq(1000) As String
Dim ArrConcAcumCod(1000) As Long
Dim ArrConcAcumCodExt(1000) As String
Dim ArrConcAcumTipo(1000) As Boolean
Dim IndConcAcum As Long

Dim Ind As Long
Dim EmpLeg As String
Dim TerApe As String
Dim TerNom As String
Dim TerNom2 As String
Dim Cuil As String
Dim FecIng As String
Dim FecBaja As String
Dim Estado As String
Dim ColActual As Long
Dim FilaActual As Long
Dim Estruc As String
Dim CausaBaja As String
Dim AnigRec As String
Dim FecNac As String
Dim EstadoCivil As String
Dim CtaSuc As String
Dim CtaNro As String
Dim Banco As String
Dim Resultado As Double
Dim cantRegistros As Long
Dim ArrayConc() As Double
Dim rs_Datos_Conc As New ADODB.Recordset

' Inicio codigo ejecutable
On Error GoTo CE

' Levanto cada parametro por separado, el separador de parametros es "@"
Flog.writeline Espacios(Tabulador * 0) & "levantando parametros" & Parametros
If Not IsNull(Parametros) Then
    
    Flog.writeline Espacios(Tabulador * 1) & "Parametros " & Parametros
    arrParam = Split(Parametros, "@")
    If UBound(arrParam) = 3 Then
    
        BuscarMonto = CBool(arrParam(0))
        ListaProc = arrParam(1)
        FecEstr = CDate(arrParam(2))
        TituloFiltro = arrParam(3)
    
        If BuscarMonto Then
            Flog.writeline Espacios(Tabulador * 1) & "Busca Monto"
            repNro = 291
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Busca Cantidad"
            repNro = 292
        End If
        Flog.writeline Espacios(Tabulador * 1) & "Procesos = " & ListaProc
        Flog.writeline Espacios(Tabulador * 1) & "Fecha Estr =" & FecEstr
        Flog.writeline Espacios(Tabulador * 1) & "Titulo = " & TituloFiltro
        
    Else
        Flog.writeline Espacios(Tabulador * 0) & "ERROR. La cantidad de parametros no es la esperada."
        HuboError = True
        Exit Sub
    End If
Else
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. No se encuentran los paramentros."
    HuboError = True
    Exit Sub
End If


Flog.writeline
'Inserto la cabecera
StrSql = "INSERT INTO repgralcab"
StrSql = StrSql & " (bpronro,titulo,fecha,monto)"
StrSql = StrSql & " VALUES("
StrSql = StrSql & " " & BproNro
StrSql = StrSql & " ,'" & TituloFiltro & "'"
StrSql = StrSql & " ," & ConvFecha(Date)
StrSql = StrSql & " ," & IIf(BuscarMonto, -1, 0)
StrSql = StrSql & " )"
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline Espacios(Tabulador * 0) & "Se inserto la cabecera del reporte."


'Configuracion del Reporte
Flog.writeline Espacios(Tabulador * 0) & "Buscando configuración del Reporte"
StrSql = "SELECT *"
StrSql = StrSql & " FROM confrep"
StrSql = StrSql & " WHERE repnro = " & repNro
StrSql = StrSql & " ORDER BY confnrocol"
OpenRecordset StrSql, rs_Confrep

If rs_Confrep.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró la configuración del Reporte " & repNro
    HuboError = True
    Exit Sub
End If

TECCosto = 0
TEConvenio = 0
TECargaHor = 0
TEPuesto = 0
TECategoria = 0
TETipoPersonal = 0
TESitoLaboral = 0
TEOSocial = 0
TEPlanOSocial = 0
TEBanco = 0
TECCostoDesc = ""
TEConvenioDesc = ""
TECargaHorDesc = ""
TEPuestoDesc = ""
TECategoriaDesc = ""
TETipoPersonalDesc = ""
TESitoLaboralDesc = ""
TEOSocialDesc = ""
TEPlanOSocialDesc = ""
TEBancoDesc = ""
IndConcAcum = 0
IndConcAcum2 = 0

Do While Not rs_Confrep.EOF
    
    Select Case CInt(rs_Confrep!confnrocol)
        Case 1:
            TECCosto = CLng(rs_Confrep!confval)
            TECCostoDesc = rs_Confrep!confetiq
        Case 2:
            TEConvenio = CLng(rs_Confrep!confval)
            TEConvenioDesc = rs_Confrep!confetiq
        Case 3:
            TECargaHor = CLng(rs_Confrep!confval)
            TECargaHorDesc = rs_Confrep!confetiq
        Case 4:
            TEPuesto = CLng(rs_Confrep!confval)
            TEPuestoDesc = rs_Confrep!confetiq
        Case 5:
            TECategoria = CLng(rs_Confrep!confval)
            TECategoriaDesc = rs_Confrep!confetiq
        Case 6:
            TETipoPersonal = CLng(rs_Confrep!confval)
            TETipoPersonalDesc = rs_Confrep!confetiq
        Case 7:
            TESitoLaboral = CLng(rs_Confrep!confval)
            TESitoLaboralDesc = rs_Confrep!confetiq
        Case 8:
            TEOSocial = CLng(rs_Confrep!confval)
            TEOSocialDesc = rs_Confrep!confetiq
        Case 9:
            TEPlanOSocial = CLng(rs_Confrep!confval)
            TEPlanOSocialDesc = rs_Confrep!confetiq
        Case 10:
            TEBanco = CLng(rs_Confrep!confval)
            TEBancoDesc = rs_Confrep!confetiq
        Case 20:
            'REMUNERATIVOS
            If Len(TipoCon1) = 0 Then
                TipoCon1 = CLng(rs_Confrep!confval)
                TipoCon1Desc = rs_Confrep!confetiq
            Else
                TipoCon1 = TipoCon1 & "," & CLng(rs_Confrep!confval)
            End If
        Case 21:
            'NO REMUNERATIVOS
            If Len(TipoCon2) = 0 Then
                TipoCon2 = CLng(rs_Confrep!confval)
                TipoCon2Desc = rs_Confrep!confetiq
            Else
                TipoCon2 = TipoCon2 & "," & CLng(rs_Confrep!confval)
            End If
        Case 22:
            'APORTES
            If Len(TipoCon3) = 0 Then
                TipoCon3 = CLng(rs_Confrep!confval)
                TipoCon3Desc = rs_Confrep!confetiq
            Else
                TipoCon3 = TipoCon3 & "," & CLng(rs_Confrep!confval)
            End If
        Case 23:
            'CONTRIBUCION
'            If BuscarMonto Then
'                If Len(TipoCon4) = 0 Then
'                    TipoCon4 = CLng(rs_Confrep!confval)
'                    TipoCon4Desc = rs_Confrep!confetiq
'                Else
'                    TipoCon4 = TipoCon4 & "," & CLng(rs_Confrep!confval)
'                End If
'            End If
            
            If BuscarMonto Then
                If UCase(rs_Confrep!conftipo) = "AC" Then
                    'Busco el acumulador
                    StrSql = "SELECT acunro, acudesabr FROM acumulador WHERE acunro = " & rs_Confrep!confval
                    OpenRecordset StrSql, rs_Consult
                    If Not rs_Consult.EOF Then
                        IndConcAcum = IndConcAcum + 1
                        ArrConcAcumEtiq(IndConcAcum) = rs_Consult!acudesabr
                        ArrConcAcumCod(IndConcAcum) = rs_Consult!acuNro
                        ArrConcAcumCodExt(IndConcAcum) = ""
                        ArrConcAcumTipo(IndConcAcum) = False
                    End If
                    rs_Consult.Close
                End If
                
                If UCase(rs_Confrep!conftipo) = "CO" Then
                    'Busco el concepto
                    StrSql = "SELECT concnro, concabr FROM concepto WHERE conccod = '" & rs_Confrep!confval2 & "'"
                    OpenRecordset StrSql, rs_Consult
                    If Not rs_Consult.EOF Then
                        IndConcAcum = IndConcAcum + 1
                        ArrConcAcumEtiq(IndConcAcum) = rs_Consult!concabr
                        ArrConcAcumCod(IndConcAcum) = rs_Consult!ConcNro
                        ArrConcAcumCodExt(IndConcAcum) = rs_Confrep!confval2
                        ArrConcAcumTipo(IndConcAcum) = True
                    End If
                    rs_Consult.Close
                End If
            End If
            
'        Case 24:
'            'PROVISIONES
'            If BuscarMonto Then
'                If Len(TipoCon5) = 0 Then
'                    TipoCon5 = CLng(rs_Confrep!confval)
'                Else
'                    TipoCon5 = TipoCon5 & "," & CLng(rs_Confrep!confval)
'                End If
'            End If
        Case 25:
            'CONCEPTOS Y ACUMULADORES
            If BuscarMonto Then
                If UCase(rs_Confrep!conftipo) = "AC" Then
                    'Busco el acumulador
                    StrSql = "SELECT acunro, acudesabr FROM acumulador WHERE acunro = " & rs_Confrep!confval
                    OpenRecordset StrSql, rs_Consult
                    If Not rs_Consult.EOF Then
                        IndConcAcum2 = IndConcAcum2 + 1
                        ArrConcAcumEtiq2(IndConcAcum2) = rs_Consult!acudesabr
                        ArrConcAcumCod2(IndConcAcum2) = rs_Consult!acuNro
                        ArrConcAcumTipo2(IndConcAcum2) = False
                    End If
                    rs_Consult.Close
                End If
                
                If UCase(rs_Confrep!conftipo) = "CO" Then
                    'Busco el concepto
                    StrSql = "SELECT concnro, concabr FROM concepto WHERE conccod = '" & rs_Confrep!confval2 & "'"
                    OpenRecordset StrSql, rs_Consult
                    If Not rs_Consult.EOF Then
                        IndConcAcum2 = IndConcAcum2 + 1
                        ArrConcAcumEtiq2(IndConcAcum2) = rs_Consult!concabr
                        ArrConcAcumCod2(IndConcAcum2) = rs_Consult!ConcNro
                        ArrConcAcumTipo2(IndConcAcum2) = True
                    End If
                    rs_Consult.Close
                End If
            End If
    End Select
    
    rs_Confrep.MoveNext
    
Loop
rs_Confrep.Close


'---------------------------------------------------------------------------------
'Inserto los titulos Fijos
'---------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Creando titulos del reporte"
Columna = 1
StrSql = "INSERT INTO repgraltit"
StrSql = StrSql & " (bpronro,descripcion,columna)"
StrSql = StrSql & " VALUES("
StrSql = StrSql & BproNro
StrSql = StrSql & " ,'Oracle ID'"
StrSql = StrSql & " ," & Columna
StrSql = StrSql & " )"
objConn.Execute StrSql, , adExecuteNoRecords
Columna = Columna + 1

StrSql = "INSERT INTO repgraltit"
StrSql = StrSql & " (bpronro,descripcion,columna)"
StrSql = StrSql & " VALUES("
StrSql = StrSql & BproNro
StrSql = StrSql & " ,'Apellidos'"
StrSql = StrSql & " ," & Columna
StrSql = StrSql & " )"
objConn.Execute StrSql, , adExecuteNoRecords
Columna = Columna + 1

StrSql = "INSERT INTO repgraltit"
StrSql = StrSql & " (bpronro,descripcion,columna)"
StrSql = StrSql & " VALUES("
StrSql = StrSql & BproNro
StrSql = StrSql & " ,'Primer Nombre Nombres'"
StrSql = StrSql & " ," & Columna
StrSql = StrSql & " )"
objConn.Execute StrSql, , adExecuteNoRecords
Columna = Columna + 1

StrSql = "INSERT INTO repgraltit"
StrSql = StrSql & " (bpronro,descripcion,columna)"
StrSql = StrSql & " VALUES("
StrSql = StrSql & BproNro
StrSql = StrSql & " ,'Segundo Nombre'"
StrSql = StrSql & " ," & Columna
StrSql = StrSql & " )"
objConn.Execute StrSql, , adExecuteNoRecords
Columna = Columna + 1

StrSql = "INSERT INTO repgraltit"
StrSql = StrSql & " (bpronro,descripcion,columna)"
StrSql = StrSql & " VALUES("
StrSql = StrSql & BproNro
StrSql = StrSql & " ,'CUIL'"
StrSql = StrSql & " ," & Columna
StrSql = StrSql & " )"
objConn.Execute StrSql, , adExecuteNoRecords
Columna = Columna + 1

StrSql = "INSERT INTO repgraltit"
StrSql = StrSql & " (bpronro,descripcion,columna)"
StrSql = StrSql & " VALUES("
StrSql = StrSql & BproNro
StrSql = StrSql & " ,'Fecha de Ingreso'"
StrSql = StrSql & " ," & Columna
StrSql = StrSql & " )"
objConn.Execute StrSql, , adExecuteNoRecords
Columna = Columna + 1

StrSql = "INSERT INTO repgraltit"
StrSql = StrSql & " (bpronro,descripcion,columna)"
StrSql = StrSql & " VALUES("
StrSql = StrSql & BproNro
StrSql = StrSql & " ,'Fecha de Baja'"
StrSql = StrSql & " ," & Columna
StrSql = StrSql & " )"
objConn.Execute StrSql, , adExecuteNoRecords
Columna = Columna + 1

StrSql = "INSERT INTO repgraltit"
StrSql = StrSql & " (bpronro,descripcion,columna)"
StrSql = StrSql & " VALUES("
StrSql = StrSql & BproNro
StrSql = StrSql & " ,'Tipo de Baja'"
StrSql = StrSql & " ," & Columna
StrSql = StrSql & " )"
objConn.Execute StrSql, , adExecuteNoRecords
Columna = Columna + 1

StrSql = "INSERT INTO repgraltit"
StrSql = StrSql & " (bpronro,descripcion,columna)"
StrSql = StrSql & " VALUES("
StrSql = StrSql & BproNro
StrSql = StrSql & " ,'Antiguedad Recon.'"
StrSql = StrSql & " ," & Columna
StrSql = StrSql & " )"
objConn.Execute StrSql, , adExecuteNoRecords
Columna = Columna + 1

If TECCosto <> 0 Then
    StrSql = "INSERT INTO repgraltit"
    StrSql = StrSql & " (bpronro,descripcion,columna)"
    StrSql = StrSql & " VALUES("
    StrSql = StrSql & BproNro
    StrSql = StrSql & " ,'" & TECCostoDesc & "'"
    StrSql = StrSql & " ," & Columna
    StrSql = StrSql & " )"
    objConn.Execute StrSql, , adExecuteNoRecords
    Columna = Columna + 1
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encuentra configurada un tipo de estructura en la columna 1 del confrep. La misma no se imprime."
End If

If TEConvenio <> 0 Then
    StrSql = "INSERT INTO repgraltit"
    StrSql = StrSql & " (bpronro,descripcion,columna)"
    StrSql = StrSql & " VALUES("
    StrSql = StrSql & BproNro
    StrSql = StrSql & " ,'" & TEConvenioDesc & "'"
    StrSql = StrSql & " ," & Columna
    StrSql = StrSql & " )"
    objConn.Execute StrSql, , adExecuteNoRecords
    Columna = Columna + 1
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encuentra configurada un tipo de estructura en la columna 2 del confrep. La misma no se imprime."
End If

If TECargaHor <> 0 Then
    StrSql = "INSERT INTO repgraltit"
    StrSql = StrSql & " (bpronro,descripcion,columna)"
    StrSql = StrSql & " VALUES("
    StrSql = StrSql & BproNro
    StrSql = StrSql & " ,'" & TECargaHorDesc & "'"
    StrSql = StrSql & " ," & Columna
    StrSql = StrSql & " )"
    objConn.Execute StrSql, , adExecuteNoRecords
    Columna = Columna + 1
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encuentra configurada un tipo de estructura en la columna 3 del confrep. La misma no se imprime."
End If

If TEPuesto <> 0 Then
    StrSql = "INSERT INTO repgraltit"
    StrSql = StrSql & " (bpronro,descripcion,columna)"
    StrSql = StrSql & " VALUES("
    StrSql = StrSql & BproNro
    StrSql = StrSql & " ,'" & TEPuestoDesc & "'"
    StrSql = StrSql & " ," & Columna
    StrSql = StrSql & " )"
    objConn.Execute StrSql, , adExecuteNoRecords
    Columna = Columna + 1
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encuentra configurada un tipo de estructura en la columna 4 del confrep. La misma no se imprime."
End If

If TECategoria <> 0 Then
    StrSql = "INSERT INTO repgraltit"
    StrSql = StrSql & " (bpronro,descripcion,columna)"
    StrSql = StrSql & " VALUES("
    StrSql = StrSql & BproNro
    StrSql = StrSql & " ,'" & TECategoriaDesc & "'"
    StrSql = StrSql & " ," & Columna
    StrSql = StrSql & " )"
    objConn.Execute StrSql, , adExecuteNoRecords
    Columna = Columna + 1
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encuentra configurada un tipo de estructura en la columna 5 del confrep. La misma no se imprime."
End If

If TETipoPersonal <> 0 Then
    StrSql = "INSERT INTO repgraltit"
    StrSql = StrSql & " (bpronro,descripcion,columna)"
    StrSql = StrSql & " VALUES("
    StrSql = StrSql & BproNro
    StrSql = StrSql & " ,'" & TETipoPersonalDesc & "'"
    StrSql = StrSql & " ," & Columna
    StrSql = StrSql & " )"
    objConn.Execute StrSql, , adExecuteNoRecords
    Columna = Columna + 1
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encuentra configurada un tipo de estructura en la columna 6 del confrep. La misma no se imprime."
End If

StrSql = "INSERT INTO repgraltit"
StrSql = StrSql & " (bpronro,descripcion,columna)"
StrSql = StrSql & " VALUES("
StrSql = StrSql & BproNro
StrSql = StrSql & " ,'Fecha Nacimiento'"
StrSql = StrSql & " ," & Columna
StrSql = StrSql & " )"
objConn.Execute StrSql, , adExecuteNoRecords
Columna = Columna + 1

StrSql = "INSERT INTO repgraltit"
StrSql = StrSql & " (bpronro,descripcion,columna)"
StrSql = StrSql & " VALUES("
StrSql = StrSql & BproNro
StrSql = StrSql & " ,'Estado'"
StrSql = StrSql & " ," & Columna
StrSql = StrSql & " )"
objConn.Execute StrSql, , adExecuteNoRecords
Columna = Columna + 1

If TESitoLaboral <> 0 Then
    StrSql = "INSERT INTO repgraltit"
    StrSql = StrSql & " (bpronro,descripcion,columna)"
    StrSql = StrSql & " VALUES("
    StrSql = StrSql & BproNro
    StrSql = StrSql & " ,'" & TESitoLaboralDesc & "'"
    StrSql = StrSql & " ," & Columna
    StrSql = StrSql & " )"
    objConn.Execute StrSql, , adExecuteNoRecords
    Columna = Columna + 1
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encuentra configurada un tipo de estructura en la columna 7 del confrep. La misma no se imprime."
End If

StrSql = "INSERT INTO repgraltit"
StrSql = StrSql & " (bpronro,descripcion,columna)"
StrSql = StrSql & " VALUES("
StrSql = StrSql & BproNro
StrSql = StrSql & " ,'Estado Civil'"
StrSql = StrSql & " ," & Columna
StrSql = StrSql & " )"
objConn.Execute StrSql, , adExecuteNoRecords
Columna = Columna + 1

If TEOSocial <> 0 Then
    StrSql = "INSERT INTO repgraltit"
    StrSql = StrSql & " (bpronro,descripcion,columna)"
    StrSql = StrSql & " VALUES("
    StrSql = StrSql & BproNro
    StrSql = StrSql & " ,'" & TEOSocialDesc & "'"
    StrSql = StrSql & " ," & Columna
    StrSql = StrSql & " )"
    objConn.Execute StrSql, , adExecuteNoRecords
    Columna = Columna + 1
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encuentra configurada un tipo de estructura en la columna 8 del confrep. La misma no se imprime."
End If

If TEPlanOSocial <> 0 Then
    StrSql = "INSERT INTO repgraltit"
    StrSql = StrSql & " (bpronro,descripcion,columna)"
    StrSql = StrSql & " VALUES("
    StrSql = StrSql & BproNro
    StrSql = StrSql & " ,'" & TEPlanOSocialDesc & "'"
    StrSql = StrSql & " ," & Columna
    StrSql = StrSql & " )"
    objConn.Execute StrSql, , adExecuteNoRecords
    Columna = Columna + 1
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encuentra configurada un tipo de estructura en la columna 9 del confrep. La misma no se imprime."
End If

'If TEBanco <> 0 Then
'    StrSql = "INSERT INTO repgraltit"
'    StrSql = StrSql & " (bpronro,descripcion,columna)"
'    StrSql = StrSql & " VALUES("
'    StrSql = StrSql & BproNro
'    StrSql = StrSql & " ,'" & TEBancoDesc & "'"
'    StrSql = StrSql & " ," & Columna
'    StrSql = StrSql & " )"
'    objConn.Execute StrSql, , adExecuteNoRecords
'    Columna = Columna + 1
'Else
'    Flog.writeline Espacios(Tabulador * 1) & "No se encuentra configurada un tipo de estructura en la columna 10 del confrep. La misma no se imprime."
'End If

StrSql = "INSERT INTO repgraltit"
StrSql = StrSql & " (bpronro,descripcion,columna)"
StrSql = StrSql & " VALUES("
StrSql = StrSql & BproNro
StrSql = StrSql & " ,'Banco'"
StrSql = StrSql & " ," & Columna
StrSql = StrSql & " )"
objConn.Execute StrSql, , adExecuteNoRecords
Columna = Columna + 1

StrSql = "INSERT INTO repgraltit"
StrSql = StrSql & " (bpronro,descripcion,columna)"
StrSql = StrSql & " VALUES("
StrSql = StrSql & BproNro
StrSql = StrSql & " ,'Sucursal Banco'"
StrSql = StrSql & " ," & Columna
StrSql = StrSql & " )"
objConn.Execute StrSql, , adExecuteNoRecords
Columna = Columna + 1

StrSql = "INSERT INTO repgraltit"
StrSql = StrSql & " (bpronro,descripcion,columna)"
StrSql = StrSql & " VALUES("
StrSql = StrSql & BproNro
StrSql = StrSql & " ,'Nro Cuenta Bank'"
StrSql = StrSql & " ," & Columna
StrSql = StrSql & " )"
objConn.Execute StrSql, , adExecuteNoRecords
Columna = Columna + 1


'---------------------------------------------------------------------------------
'Busco los tipos de conceptos remunerativos
'---------------------------------------------------------------------------------
IndRem = 0
If Len(TipoCon1) <> 0 Then
    StrSql = "SELECT concnro, conccod, concabr, tconnro FROM concepto"
    StrSql = StrSql & " Where tconnro IN (" & TipoCon1 & ")"
    StrSql = StrSql & " AND concimp = -1"
    StrSql = StrSql & " ORDER BY tconnro ,conccod"
    OpenRecordset StrSql, rs_Consult
    Do While Not rs_Consult.EOF
        If IndRem < 1000 Then
            IndRem = IndRem + 1
            ArrRemEtiq(IndRem) = rs_Consult!concabr
            ArrRemCod(IndRem) = rs_Consult!ConcNro
            ArrRemTipo(IndRem) = rs_Consult!tconnro
            StrSql = "INSERT INTO repgraltit"
            StrSql = StrSql & " (bpronro,descripcion,columna)"
            StrSql = StrSql & " VALUES("
            StrSql = StrSql & BproNro
            StrSql = StrSql & " ,'" & rs_Consult!concabr & "'"
            StrSql = StrSql & " ," & Columna
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Columna = Columna + 1
        End If
        
        rs_Consult.MoveNext
    Loop
    rs_Consult.Close
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encuentra configurada un tipo de concepto en la columna 20 del confrep. La misma no se imprime."
End If

'Creo la columna de total
If IndRem > 0 Then
    StrSql = "INSERT INTO repgraltit"
    StrSql = StrSql & " (bpronro,descripcion,columna)"
    StrSql = StrSql & " VALUES("
    StrSql = StrSql & BproNro
    StrSql = StrSql & " ,'Total " & TipoCon1Desc & "'"
    StrSql = StrSql & " ," & Columna
    StrSql = StrSql & " )"
    objConn.Execute StrSql, , adExecuteNoRecords
    Columna = Columna + 1
End If


'---------------------------------------------------------------------------------
'Busco los tipos de conceptos no remunerativos
'---------------------------------------------------------------------------------
IndNoRem = 0
If Len(TipoCon2) <> 0 Then
    StrSql = "SELECT concnro, conccod, concabr FROM concepto"
    StrSql = StrSql & " Where tconnro IN (" & TipoCon2 & ")"
    StrSql = StrSql & " AND concimp = -1"
    StrSql = StrSql & " ORDER BY tconnro ,conccod"
    OpenRecordset StrSql, rs_Consult
    Do While Not rs_Consult.EOF
        
        If IndNoRem < 1000 Then
            IndNoRem = IndNoRem + 1
            ArrNoRemEtiq(IndNoRem) = rs_Consult!concabr
            ArrNoRemCod(IndNoRem) = rs_Consult!ConcNro
            
            StrSql = "INSERT INTO repgraltit"
            StrSql = StrSql & " (bpronro,descripcion,columna)"
            StrSql = StrSql & " VALUES("
            StrSql = StrSql & BproNro
            StrSql = StrSql & " ,'" & rs_Consult!concabr & "'"
            StrSql = StrSql & " ," & Columna
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
            Columna = Columna + 1
            
        End If
        
        rs_Consult.MoveNext
    Loop
    rs_Consult.Close
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encuentra configurada un tipo de concepto en la columna 21 del confrep. La misma no se imprime."
End If

'Creo la columna de total
If IndNoRem > 0 Then
    StrSql = "INSERT INTO repgraltit"
    StrSql = StrSql & " (bpronro,descripcion,columna)"
    StrSql = StrSql & " VALUES("
    StrSql = StrSql & BproNro
    StrSql = StrSql & " ,'Total " & TipoCon2Desc & "'"
    StrSql = StrSql & " ," & Columna
    StrSql = StrSql & " )"
    objConn.Execute StrSql, , adExecuteNoRecords
    Columna = Columna + 1
End If


'---------------------------------------------------------------------------------
'Busco los tipos de conceptos de retenciones
'---------------------------------------------------------------------------------
IndRet = 0
If Len(TipoCon3) <> 0 Then
    StrSql = "SELECT concnro, conccod, concabr FROM concepto"
    StrSql = StrSql & " Where tconnro IN (" & TipoCon3 & ")"
    StrSql = StrSql & " AND concimp = -1"
    StrSql = StrSql & " ORDER BY tconnro ,conccod"
    OpenRecordset StrSql, rs_Consult
    Do While Not rs_Consult.EOF
        
        If IndRet < 1000 Then
            IndRet = IndRet + 1
            ArrRetEtiq(IndRet) = rs_Consult!concabr
            ArrRetCod(IndRet) = rs_Consult!ConcNro
            
            StrSql = "INSERT INTO repgraltit"
            StrSql = StrSql & " (bpronro,descripcion,columna)"
            StrSql = StrSql & " VALUES("
            StrSql = StrSql & BproNro
            StrSql = StrSql & " ,'" & rs_Consult!concabr & "'"
            StrSql = StrSql & " ," & Columna
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
            Columna = Columna + 1
            
        End If
        
        rs_Consult.MoveNext
    Loop
    rs_Consult.Close
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encuentra configurada un tipo de concepto en la columna 22 del confrep. La misma no se imprime."
End If

'Creo la columna de total
If IndRet > 0 Then
    StrSql = "INSERT INTO repgraltit"
    StrSql = StrSql & " (bpronro,descripcion,columna)"
    StrSql = StrSql & " VALUES("
    StrSql = StrSql & BproNro
    StrSql = StrSql & " ,'Total " & TipoCon3Desc & "'"
    StrSql = StrSql & " ," & Columna
    StrSql = StrSql & " )"
    objConn.Execute StrSql, , adExecuteNoRecords
    Columna = Columna + 1
End If


If BuscarMonto Then
    '---------------------------------------------------------------------------------
    'Creo la columna Sueldo Neto
    '---------------------------------------------------------------------------------
    If ((IndNoRem > 0) Or (IndRem > 0) Or (IndRet > 0)) Then
        StrSql = "INSERT INTO repgraltit"
        StrSql = StrSql & " (bpronro,descripcion,columna)"
        StrSql = StrSql & " VALUES("
        StrSql = StrSql & BproNro
        StrSql = StrSql & " ,'Sueldos Neto'"
        StrSql = StrSql & " ," & Columna
        StrSql = StrSql & " )"
        objConn.Execute StrSql, , adExecuteNoRecords
        Columna = Columna + 1
    End If

    '---------------------------------------------------------------------------------
    'Busco los tipos de conceptos de contribuciones
    '---------------------------------------------------------------------------------
'    IndContr = 0
'    If Len(TipoCon4) <> 0 Then
'        StrSql = "SELECT concnro, conccod, concabr FROM concepto"
'        'StrSql = StrSql & " Where tconnro IN (" & TipoCon2 & ")"
'        StrSql = StrSql & " Where tconnro IN (" & TipoCon4 & ")"
'        StrSql = StrSql & " AND concimp <> -1"
'        StrSql = StrSql & " ORDER BY tconnro ,conccod"
'        OpenRecordset StrSql, rs_Consult
'        Do While Not rs_Consult.EOF
'
'            If IndRem < 1000 Then
'                IndContr = IndContr + 1
'                ArrContrEtiq(IndContr) = rs_Consult!concabr
'                ArrContrCod(IndContr) = rs_Consult!concnro
'
'                StrSql = "INSERT INTO repgraltit"
'                StrSql = StrSql & " (bpronro,descripcion,columna)"
'                StrSql = StrSql & " VALUES("
'                StrSql = StrSql & BproNro
'                StrSql = StrSql & " ,'" & rs_Consult!concabr & "'"
'                StrSql = StrSql & " ," & Columna
'                StrSql = StrSql & " )"
'                objConn.Execute StrSql, , adExecuteNoRecords
'                Columna = Columna + 1
'
'            End If
'
'            rs_Consult.MoveNext
'        Loop
'        rs_Consult.Close
'    Else
'        Flog.writeline Espacios(Tabulador * 1) & "No se encuentra configurada un tipo de concepto en la columna 23 del confrep. La misma no se imprime."
'    End If
    
    'Creo la columna de total
'    If IndContr > 0 Then
'        StrSql = "INSERT INTO repgraltit"
'        StrSql = StrSql & " (bpronro,descripcion,columna)"
'        StrSql = StrSql & " VALUES("
'        StrSql = StrSql & BproNro
'        StrSql = StrSql & " ,'Total " & TipoCon4Desc & "'"
'        StrSql = StrSql & " ," & Columna
'        StrSql = StrSql & " )"
'        objConn.Execute StrSql, , adExecuteNoRecords
'        Columna = Columna + 1
'    End If
    If BuscarMonto Then
        If IndConcAcum = 0 Then Flog.writeline Espacios(Tabulador * 1) & "No se encuentra configurada un conceptos y acumuladores en la columna 23 del confrep. La misma no se imprime."
        
        For Ind = 1 To IndConcAcum
            If Ind < 1000 Then
                StrSql = "INSERT INTO repgraltit"
                StrSql = StrSql & " (bpronro,descripcion,columna)"
                StrSql = StrSql & " VALUES("
                StrSql = StrSql & BproNro
                StrSql = StrSql & " ,'" & ArrConcAcumEtiq(Ind) & "'"
                StrSql = StrSql & " ," & Columna
                StrSql = StrSql & " )"
                objConn.Execute StrSql, , adExecuteNoRecords
                Columna = Columna + 1
            End If
        Next

        'Creo la columna de total
        If IndConcAcum > 0 Then
            StrSql = "INSERT INTO repgraltit"
            StrSql = StrSql & " (bpronro,descripcion,columna)"
            StrSql = StrSql & " VALUES("
            StrSql = StrSql & BproNro
            StrSql = StrSql & " ,'Total CCSS'"
            StrSql = StrSql & " ," & Columna
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
            Columna = Columna + 1
        End If
    
        '---------------------------------------------------------------------------------
        'Busco los tipos de conceptos de Provisiones
        '---------------------------------------------------------------------------------
    '    IndProv = 0
    '    If Len(TipoCon5) <> 0 Then
    '        StrSql = "SELECT concnro, conccod, concabr FROM concepto"
    '        'StrSql = StrSql & " Where tconnro IN (" & TipoCon2 & ")"
    '        StrSql = StrSql & " Where tconnro IN (" & TipoCon5 & ")"
    '        StrSql = StrSql & " AND concimp <> -1"
    '        StrSql = StrSql & " ORDER BY tconnro ,conccod"
    '        OpenRecordset StrSql, rs_Consult
    '        Do While Not rs_Consult.EOF
    '
    '            If IndRem < 1000 Then
    '                IndProv = IndProv + 1
    '                ArrProvEtiq(IndProv) = rs_Consult!concabr
    '                ArrProvCod(IndProv) = rs_Consult!concnro
    '
    '                StrSql = "INSERT INTO repgraltit"
    '                StrSql = StrSql & " (bpronro,descripcion,columna)"
    '                StrSql = StrSql & " VALUES("
    '                StrSql = StrSql & BproNro
    '                StrSql = StrSql & " ,'" & rs_Consult!concabr & "'"
    '                StrSql = StrSql & " ," & Columna
    '                StrSql = StrSql & " )"
    '                objConn.Execute StrSql, , adExecuteNoRecords
    '                Columna = Columna + 1
    '
    '            End If
    '
    '            rs_Consult.MoveNext
    '        Loop
    '        rs_Consult.Close
    '    Else
    '        Flog.writeline Espacios(Tabulador * 1) & "No se encuentra configurada un tipo de concepto en la columna 24 del confrep. La misma no se imprime."
    '    End If
    
    
        '---------------------------------------------------------------------------------
        'Busco conceptos y acumuladores
        '---------------------------------------------------------------------------------
        If IndConcAcum2 = 0 Then Flog.writeline Espacios(Tabulador * 1) & "No se encuentra configurada un conceptos y acumuladores en la columna 25 del confrep. La misma no se imprime."
        
        For Ind = 1 To IndConcAcum2
            If Ind < 1000 Then
                StrSql = "INSERT INTO repgraltit"
                StrSql = StrSql & " (bpronro,descripcion,columna)"
                StrSql = StrSql & " VALUES("
                StrSql = StrSql & BproNro
                StrSql = StrSql & " ,'" & ArrConcAcumEtiq2(Ind) & "'"
                StrSql = StrSql & " ," & Columna
                StrSql = StrSql & " )"
                objConn.Execute StrSql, , adExecuteNoRecords
                Columna = Columna + 1
            End If
        Next
    
    End If
    '---------------------------------------------------------------------------------
    'Total Costo Empresa
    '---------------------------------------------------------------------------------
    'If ((IndNoRem > 0) Or (IndRem > 0) Or (IndRet > 0) Or (IndContr > 0) Or (IndProv > 0) Or (IndConcAcum2 > 0)) Then
    If ((IndNoRem > 0) Or (IndRem > 0) Or (IndRet > 0) Or (IndConcAcum > 0) Or (IndConcAcum2 > 0)) Then
        StrSql = "INSERT INTO repgraltit"
        StrSql = StrSql & " (bpronro,descripcion,columna)"
        StrSql = StrSql & " VALUES("
        StrSql = StrSql & BproNro
        StrSql = StrSql & " ,'Total Costo Empresa'"
        StrSql = StrSql & " ," & Columna
        StrSql = StrSql & " )"
        objConn.Execute StrSql, , adExecuteNoRecords
        Columna = Columna + 1
    End If
    
    Flog.writeline Espacios(Tabulador * 1) & "Fin de busqueda de titulos de las columnas"
    Flog.writeline
    
End If 'If BuscarMonto Then

'---------------------------------------------------------------------------------
'Consulta Principal
'---------------------------------------------------------------------------------
StrSql = "SELECT *"
StrSql = StrSql & " FROM batch_empleado"
StrSql = StrSql & " WHERE bpronro = " & BproNro
StrSql = StrSql & " ORDER BY progreso"
OpenRecordset StrSql, rs_Empleados

'seteo de las variables de progreso
Progreso = 0
CEmpleadosAProc = rs_Empleados.RecordCount
cantRegistros = CEmpleadosAProc

'---------------------------------------------------------------------------------
'Actualizo el progreso
'---------------------------------------------------------------------------------
TiempoAcumulado = GetTickCount
StrSql = "UPDATE batch_proceso SET bprcprogreso = 10" & _
         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
         "', bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProcesoBatch
objconnProgreso.Execute StrSql, , adExecuteNoRecords


If CEmpleadosAProc = 0 Then
   Flog.writeline "no hay empleados"
   CEmpleadosAProc = 1
Else
    Flog.writeline
    Flog.writeline
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "------------------------------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 0) & "Comienza el procesamiento de " & CEmpleadosAProc & " empleados."
    Flog.writeline Espacios(Tabulador * 0) & "------------------------------------------------------------------------"
End If
IncPorc = (90 / CEmpleadosAProc)
        
FilaActual = 0

'Comienzo a procesar los empleados
Do While Not rs_Empleados.EOF
    '----------------------------------------------------------------------
    'GUARDO TODOS LOS MONTOS/CANTIDADES DEL DETLIQ DEL EMPLEADO EN UN ARRAY
    '----------------------------------------------------------------------
    ReDim ArrayConc(99999) As Double
    
    If BuscarMonto Then
        StrSql = "SELECT dlimonto guarda, concnro"
    Else
        StrSql = "SELECT detliq.dlicant guarda, concnro"
    End If
    StrSql = StrSql & " FROM cabliq"
    StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro"
    StrSql = StrSql & " WHERE Empleado = " & rs_Empleados!Ternro
    StrSql = StrSql & " AND pronro IN (" & ListaProc & ")"
    OpenRecordset StrSql, rs_Datos_Conc
    Do While Not rs_Datos_Conc.EOF
        'Flog.writeline "concnro: " & rs_Datos_Conc!concnro
        ArrayConc(CLng(rs_Datos_Conc!ConcNro)) = ArrayConc(CLng(rs_Datos_Conc!ConcNro)) + CDbl(rs_Datos_Conc!guarda)
        
        rs_Datos_Conc.MoveNext
    Loop
    rs_Datos_Conc.Close
    
    Set rs_Datos_Conc = Nothing
    '----------------------------------------------------------------------
    '----------------------------------------------------------------------
    
    FilaActual = FilaActual + 1
    ColActual = 0
    
    'Busco los datos del empleado
    '---------------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 1) & "Buscando datos del empleado ternro " & rs_Empleados!Ternro
    EmpLeg = ""
    TerApe = ""
    TerNom = ""
    TerNom2 = ""
    Estado = ""
    FecNac = ""
    EstadoCivil = ""
    StrSql = "SELECT empleado.empleg, empleado.terape, empleado.ternom, tercero.terfecnac,"
    StrSql = StrSql & " empleado.terape2, empleado.ternom2, empleado.empest, estcivdesabr"
    StrSql = StrSql & " FROM empleado"
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro"
    StrSql = StrSql & " LEFT JOIN estcivil ON tercero.estcivnro = estcivil.estcivnro"
    StrSql = StrSql & " WHERE empleado.ternro = " & rs_Empleados!Ternro
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        EmpLeg = rs_Consult!EmpLeg
        TerApe = rs_Consult!TerApe
        If Not EsNulo(rs_Consult!TerApe2) Then TerApe = TerApe & " " & rs_Consult!TerApe2
        TerNom = rs_Consult!TerNom
        TerNom2 = IIf(EsNulo(rs_Consult!TerNom2), "", rs_Consult!TerNom2)
        Estado = IIf(CBool(rs_Consult!empest), "Activo", "Baja")
        FecNac = IIf(EsNulo(rs_Consult!terfecnac), "", rs_Consult!terfecnac)
        EstadoCivil = IIf(EsNulo(rs_Consult!estcivdesabr), "", rs_Consult!estcivdesabr)
    End If
    rs_Consult.Close
    
    ColActual = ColActual + 1
    Call insertarDet(BproNro, rs_Empleados!Ternro, EmpLeg, ColActual, FilaActual)
    
    ColActual = ColActual + 1
    Call insertarDet(BproNro, rs_Empleados!Ternro, TerApe, ColActual, FilaActual)
    
    ColActual = ColActual + 1
    Call insertarDet(BproNro, rs_Empleados!Ternro, TerNom, ColActual, FilaActual)
    
    ColActual = ColActual + 1
    Call insertarDet(BproNro, rs_Empleados!Ternro, TerNom2, ColActual, FilaActual)
    
    'Busco CUIL del empleado
    '---------------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 1) & "Buscando CUIL"
    Cuil = ""
    StrSql = "SELECT nrodoc"
    StrSql = StrSql & " FROM ter_doc"
    StrSql = StrSql & " WHERE tidnro = 10"
    StrSql = StrSql & " AND ternro = " & rs_Empleados!Ternro
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then Cuil = rs_Consult!Nrodoc
    rs_Consult.Close
    
    ColActual = ColActual + 1
    Call insertarDet(BproNro, rs_Empleados!Ternro, Cuil, ColActual, FilaActual)
    
    'Busco Fases del empleado
    '---------------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 1) & "Fases"
    FecIng = ""
    FecBaja = ""
    CausaBaja = ""
'    StrSql = "SELECT altfec, bajfec, caudes FROM fases"
'    StrSql = StrSql & " LEFT JOIN causa ON causa.caunro = fases.caunro"
'    StrSql = StrSql & " WHERE empleado = " & rs_Empleados!ternro
'    StrSql = StrSql & " AND altfec <= " & ConvFecha(FecEstr)
'    StrSql = StrSql & " AND ((bajfec > " & ConvFecha(FecEstr) & ") OR (bajfec IS NULL))"
    StrSql = "SELECT altfec, bajfec, caudes FROM fases"
    StrSql = StrSql & " LEFT JOIN causa ON causa.caunro = fases.caunro"
    StrSql = StrSql & " WHERE empleado = " & rs_Empleados!Ternro
    StrSql = StrSql & " AND altfec <= " & ConvFecha(FecEstr)
    StrSql = StrSql & " ORDER BY altfec DESC "
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        FecIng = rs_Consult!altfec
        FecBaja = IIf(EsNulo(rs_Consult!bajfec), "", rs_Consult!bajfec)
        CausaBaja = IIf(EsNulo(rs_Consult!caudes), "", rs_Consult!caudes)
    End If
    rs_Consult.Close
    
    ColActual = ColActual + 1
    Call insertarDet(BproNro, rs_Empleados!Ternro, FecIng, ColActual, FilaActual)
    
    ColActual = ColActual + 1
    Call insertarDet(BproNro, rs_Empleados!Ternro, FecBaja, ColActual, FilaActual)
    
    ColActual = ColActual + 1
    Call insertarDet(BproNro, rs_Empleados!Ternro, CausaBaja, ColActual, FilaActual)
    
    'Busco Antig Reconocida
    '---------------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 1) & "Antiguedad"
    AnigRec = ""
    StrSql = "SELECT altfec FROM fases"
    StrSql = StrSql & " WHERE empleado = " & rs_Empleados!Ternro
    StrSql = StrSql & " AND fasrecofec = -1"
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then AnigRec = rs_Consult!altfec
    rs_Consult.Close
    
    ColActual = ColActual + 1
    Call insertarDet(BproNro, rs_Empleados!Ternro, AnigRec, ColActual, FilaActual)
    
    If TECCosto <> 0 Then
        Call buscarEstr(rs_Empleados!Ternro, TECCosto, FecEstr, Estruc)
        ColActual = ColActual + 1
        Call insertarDet(BproNro, rs_Empleados!Ternro, Estruc, ColActual, FilaActual)
    End If
    
    If TEConvenio <> 0 Then
        Call buscarEstr(rs_Empleados!Ternro, TEConvenio, FecEstr, Estruc)
        ColActual = ColActual + 1
        Call insertarDet(BproNro, rs_Empleados!Ternro, Estruc, ColActual, FilaActual)
    End If
    
    If TECargaHor <> 0 Then
        Call buscarEstr(rs_Empleados!Ternro, TECargaHor, FecEstr, Estruc)
        ColActual = ColActual + 1
        Call insertarDet(BproNro, rs_Empleados!Ternro, Estruc, ColActual, FilaActual)
    End If
    
    If TEPuesto <> 0 Then
        Call buscarEstr(rs_Empleados!Ternro, TEPuesto, FecEstr, Estruc)
        ColActual = ColActual + 1
        Call insertarDet(BproNro, rs_Empleados!Ternro, Estruc, ColActual, FilaActual)
    End If
    
    If TECategoria <> 0 Then
        Call buscarEstr(rs_Empleados!Ternro, TECategoria, FecEstr, Estruc)
        ColActual = ColActual + 1
        Call insertarDet(BproNro, rs_Empleados!Ternro, Estruc, ColActual, FilaActual)
    End If
    
    If TETipoPersonal <> 0 Then
        Call buscarEstr(rs_Empleados!Ternro, TETipoPersonal, FecEstr, Estruc)
        ColActual = ColActual + 1
        Call insertarDet(BproNro, rs_Empleados!Ternro, Estruc, ColActual, FilaActual)
    End If

    ColActual = ColActual + 1
    Call insertarDet(BproNro, rs_Empleados!Ternro, FecNac, ColActual, FilaActual)
    
    ColActual = ColActual + 1
    Call insertarDet(BproNro, rs_Empleados!Ternro, Estado, ColActual, FilaActual)
    
    If TESitoLaboral <> 0 Then
        Call buscarEstr(rs_Empleados!Ternro, TESitoLaboral, FecEstr, Estruc)
        ColActual = ColActual + 1
        Call insertarDet(BproNro, rs_Empleados!Ternro, Estruc, ColActual, FilaActual)
    End If

    ColActual = ColActual + 1
    Call insertarDet(BproNro, rs_Empleados!Ternro, EstadoCivil, ColActual, FilaActual)

    If TEOSocial <> 0 Then
        Call buscarEstr(rs_Empleados!Ternro, TEOSocial, FecEstr, Estruc)
        ColActual = ColActual + 1
        Call insertarDet(BproNro, rs_Empleados!Ternro, Estruc, ColActual, FilaActual)
    End If

    If TEPlanOSocial <> 0 Then
        Call buscarEstr(rs_Empleados!Ternro, TEPlanOSocial, FecEstr, Estruc)
        ColActual = ColActual + 1
        Call insertarDet(BproNro, rs_Empleados!Ternro, Estruc, ColActual, FilaActual)
    End If

'    If TEBanco <> 0 Then
'        Call buscarEstr(rs_Empleados!ternro, TEBanco, FecEstr, Estruc)
'        ColActual = ColActual + 1
'        Call insertarDet(BproNro, rs_Empleados!ternro, Estruc, ColActual, FilaActual)
'    End If

    '---------------------------------------------------------------------------------
    'Datos de la cuenta bancaria
    '---------------------------------------------------------------------------------
    '10/07/2014 - FB - Se corrige el mensaje del log
    'Flog.writeline Espacios(Tabulador * 1) & "Buscando CUIL"
    Flog.writeline Espacios(Tabulador * 1) & "Buscando Datos de la Cta. Bancaria"
    CtaSuc = ""
    CtaNro = ""
    Banco = ""
    StrSql = "SELECT ctabancaria.ctabnro, ctabancaria.ctabsuc, banco.bandesc FROM ctabancaria"
    StrSql = StrSql & " INNER JOIN banco ON banco.ternro = ctabancaria.banco"
    StrSql = StrSql & " WHERE ctabancaria.ternro = " & rs_Empleados!Ternro
    StrSql = StrSql & " AND ctabestado = -1"
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        CtaSuc = IIf(EsNulo(rs_Consult!ctabsuc), "", rs_Consult!ctabsuc)
        CtaNro = IIf(EsNulo(rs_Consult!ctabnro), "", rs_Consult!ctabnro)
        Banco = IIf(EsNulo(rs_Consult!bandesc), "", rs_Consult!bandesc)
    End If
    rs_Consult.Close
    
    ColActual = ColActual + 1
    Call insertarDet(BproNro, rs_Empleados!Ternro, Banco, ColActual, FilaActual)
    
    ColActual = ColActual + 1
    Call insertarDet(BproNro, rs_Empleados!Ternro, CtaSuc, ColActual, FilaActual)
    
    ColActual = ColActual + 1
    '10/07/2014 - FB - Se agrega un "espacio" delante de la cuenta bancaria para la exportacion a excel como texto.
    'Call insertarDet(BproNro, rs_Empleados!Ternro, CtaNro, ColActual, FilaActual)
    Call insertarDet(BproNro, rs_Empleados!Ternro, " " & CtaNro, ColActual, FilaActual)
    
    '---------------------------------------------------------------------------------
    'Conceptos Remunerativos
    '---------------------------------------------------------------------------------
    TotalRem = 0
    TotalRemEmpresa = 0
    For Ind = 1 To IndRem
        'Call buscarConc(rs_Empleados!Ternro, ArrRemCod(Ind), ListaProc, BuscarMonto, Resultado)
        Resultado = ArrayConc(ArrRemCod(Ind))
        
        ColActual = ColActual + 1
        
        If ((ArrRemTipo(Ind) <> 2) And (ArrRemTipo(Ind) <> 2)) Then TotalRemEmpresa = TotalRemEmpresa + Resultado
        TotalRem = TotalRem + Resultado
        
        Call insertarDet(BproNro, rs_Empleados!Ternro, Format(Resultado, "Fixed"), ColActual, FilaActual)
    Next
    If IndRem > 0 Then
        ColActual = ColActual + 1
        Call insertarDet(BproNro, rs_Empleados!Ternro, Format(TotalRem, "Fixed"), ColActual, FilaActual)
    End If
    
    '---------------------------------------------------------------------------------
    'Conceptos No Remunerativos
    '---------------------------------------------------------------------------------
    TotalNoRem = 0
    For Ind = 1 To IndNoRem
        'Call buscarConc(rs_Empleados!Ternro, ArrNoRemCod(Ind), ListaProc, BuscarMonto, Resultado)
        
        Resultado = ArrayConc(ArrNoRemCod(Ind))
        
        ColActual = ColActual + 1
        TotalNoRem = TotalNoRem + Resultado
        Call insertarDet(BproNro, rs_Empleados!Ternro, Format(Resultado, "Fixed"), ColActual, FilaActual)
    Next
    If IndNoRem > 0 Then
        ColActual = ColActual + 1
        Call insertarDet(BproNro, rs_Empleados!Ternro, Format(TotalNoRem, "Fixed"), ColActual, FilaActual)
    End If
    
    '---------------------------------------------------------------------------------
    'Conceptos Retenciones
    '---------------------------------------------------------------------------------
    TotalRet = 0
    For Ind = 1 To IndRet
        'Call buscarConc(rs_Empleados!Ternro, ArrRetCod(Ind), ListaProc, BuscarMonto, Resultado)
        
        Resultado = ArrayConc(ArrRetCod(Ind))
        
        ColActual = ColActual + 1
        TotalRet = TotalRet + Resultado
        Call insertarDet(BproNro, rs_Empleados!Ternro, Format(Resultado, "Fixed"), ColActual, FilaActual)
    Next
    If IndRet > 0 Then
        ColActual = ColActual + 1
        Call insertarDet(BproNro, rs_Empleados!Ternro, Format(TotalRet, "Fixed"), ColActual, FilaActual)
    End If
    
    
    If BuscarMonto Then
        '---------------------------------------------------------------------------------
        'Sueldos Neto
        '---------------------------------------------------------------------------------
        If ((IndNoRem > 0) Or (IndRem > 0) Or (IndRet > 0)) Then
            ColActual = ColActual + 1
            Call insertarDet(BproNro, rs_Empleados!Ternro, Format(TotalRem + TotalNoRem + TotalRet, "Fixed"), ColActual, FilaActual)
        End If
    
        '---------------------------------------------------------------------------------
        'Conceptos Contribuciones
        '---------------------------------------------------------------------------------
        TotalContr = 0
        TotalContrEmpresa = 0
'       For Ind = 1 To IndContr
'           Call buscarConc(rs_Empleados!ternro, ArrContrCod(Ind), ListaProc, BuscarMonto, Resultado)
'           ColActual = ColActual + 1
'           TotalContr = TotalContr + Resultado
'           Call insertarDet(BproNro, rs_Empleados!ternro, Format(Resultado, "Fixed"), ColActual, FilaActual)
'       Next
'       If IndContr > 0 Then
'           ColActual = ColActual + 1
'           Call insertarDet(BproNro, rs_Empleados!ternro, Format(TotalContr, "Fixed"), ColActual, FilaActual)
'       End If
        For Ind = 1 To IndConcAcum
            If ArrConcAcumTipo(Ind) Then
                'Call buscarConc(rs_Empleados!Ternro, ArrConcAcumCod(Ind), ListaProc, BuscarMonto, Resultado)
                
                Resultado = ArrayConc(ArrConcAcumCod(Ind))
            Else
                Call buscarAcum(rs_Empleados!Ternro, ArrConcAcumCod(Ind), ListaProc, BuscarMonto, Resultado)
            End If
            ColActual = ColActual + 1
            If ((ArrConcAcumCodExt(Ind) <> "11330") And (ArrConcAcumCodExt(Ind) <> "11340")) Then
                TotalContr = TotalContr + Abs(Resultado)
            End If
            Call insertarDet(BproNro, rs_Empleados!Ternro, Format(Abs(Resultado), "Fixed"), ColActual, FilaActual)
            'If ((ArrConcAcumCodExt(Ind) <> "11330") And (ArrConcAcumCodExt(Ind) <> "11340")) Then
                TotalContrEmpresa = TotalContrEmpresa + Abs(Resultado)
            'End If
        Next
        If IndConcAcum > 0 Then
            ColActual = ColActual + 1
            Call insertarDet(BproNro, rs_Empleados!Ternro, Format(TotalContr, "Fixed"), ColActual, FilaActual)
        End If
        
        '---------------------------------------------------------------------------------
        'Conceptos Provisiones y conceptos y acumuladores configurados
        '---------------------------------------------------------------------------------
        TotalProv = 0
'       For Ind = 1 To IndProv
'           Call buscarConc(rs_Empleados!ternro, ArrProvCod(Ind), ListaProc, BuscarMonto, Resultado)
'           ColActual = ColActual + 1
'           TotalProv = TotalProv + Resultado
'           Call insertarDet(BproNro, rs_Empleados!ternro, Format(Resultado, "Fixed"), ColActual, FilaActual)
'       Next
        
        
        For Ind = 1 To IndConcAcum2
            If ArrConcAcumTipo2(Ind) Then
                'Call buscarConc(rs_Empleados!Ternro, ArrConcAcumCod2(Ind), ListaProc, BuscarMonto, Resultado)
                
                Resultado = ArrayConc(ArrConcAcumCod2(Ind))
            Else
                Call buscarAcum(rs_Empleados!Ternro, ArrConcAcumCod2(Ind), ListaProc, BuscarMonto, Resultado)
            End If
            ColActual = ColActual + 1
            TotalProv = TotalProv + Abs(Resultado)
            Call insertarDet(BproNro, rs_Empleados!Ternro, Format(Abs(Resultado), "Fixed"), ColActual, FilaActual)
        Next
        
        
        '---------------------------------------------------------------------------------
        'Total por empresa
        '---------------------------------------------------------------------------------
        'If ((IndNoRem > 0) Or (IndRem > 0) Or (IndRet > 0) Or (IndContr > 0) Or (IndProv > 0) Or (IndConcAcum2 > 0)) Then
        If ((IndNoRem > 0) Or (IndRem > 0) Or (IndRet > 0) Or (IndConcAcum > 0) Or (IndConcAcum2 > 0)) Then
            ColActual = ColActual + 1
            'Call insertarDet(BproNro, rs_Empleados!ternro, Format(TotalRem + TotalNoRem + TotalRet + TotalContr + TotalProv, "Fixed"), ColActual, FilaActual)
            Call insertarDet(BproNro, rs_Empleados!Ternro, Format(TotalRemEmpresa + TotalNoRem + TotalContrEmpresa + TotalProv, "Fixed"), ColActual, FilaActual)
        End If
    End If 'If BuscarMonto Then
    
    
    
    '---------------------------------------------------------------------------------
    'Actualizo el progreso
    '---------------------------------------------------------------------------------
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    cantRegistros = cantRegistros - 1
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "', bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Paso a siguiente cabliq
    rs_Empleados.MoveNext
    
Loop



If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_Consult.State = adStateOpen Then rs_Consult.Close


Set rs_Empleados = Nothing
Set rs_Confrep = Nothing
Set rs_Consult = Nothing

Exit Sub

CE:
    Flog.writeline "=================================================================="
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

End Sub


Public Sub buscarConc(ByVal ConcNro As Long, ByRef ArrayDeConc As Variant, ByRef Valor As Double)
'Public Sub buscarConc(ByVal Ternro As Long, ByVal concnro As Long, ByVal ListaProc As String, ByVal MontoCant As Boolean, ByRef Valor As Double)
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca Monto o cant liquidado para un concepto para un empleado y una lista de procesos.
' Autor      : Martin Ferraro
' Fecha      : 01/10/2010
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Dim rs_Datos As New ADODB.Recordset
'Dim Monto As Double
'Dim cant As Double

    'Monto = 0
    'cant = 0
    
    'StrSql = "SELECT SUM(dlimonto) monto, SUM(detliq.dlicant) cant"
    'StrSql = StrSql & " FROM cabliq"
    'StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro AND detliq.concnro = " & concnro
    'StrSql = StrSql & " WHERE Empleado = " & Ternro
    'StrSql = StrSql & " AND pronro IN (" & ListaProc & ")"
    ''StrSql = StrSql & " AND detliq.concnro = " & concnro
    'OpenRecordset StrSql, rs_Datos
    'If Not rs_Datos.EOF Then
        'Monto = IIf(EsNulo(rs_Datos!Monto), 0, rs_Datos!Monto)
        'cant = IIf(EsNulo(rs_Datos!cant), 0, rs_Datos!cant)
    'End If
    
    'rs_Datos.Close
    
    'If MontoCant Then
        'Valor = Monto
    'Else
        'Valor = cant
    'End If
    
    Valor = ArrayDeConc(ConcNro)
    
'Set rs_Datos = Nothing
End Sub


Public Sub buscarAcum(ByVal Ternro As Long, ByVal acuNro As Long, ByVal ListaProc As String, ByVal MontoCant As Boolean, ByRef Valor As Double)
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca Monto o cant liquidado para un acumulador para un empleado y una lista de procesos.
' Autor      : Martin Ferraro
' Fecha      : 01/10/2010
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Datos As New ADODB.Recordset
Dim Monto As Double
Dim cant As Double

    Monto = 0
    cant = 0
    
    StrSql = "SELECT SUM(almonto) monto, SUM(acu_liq.alcant) cant"
    StrSql = StrSql & " FROM cabliq"
    StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro AND acu_liq.acunro = " & acuNro
    StrSql = StrSql & " WHERE empleado = " & Ternro
    StrSql = StrSql & " AND pronro IN (" & ListaProc & ")"
    'StrSql = StrSql & " AND acu_liq.acunro = " & acuNro
    OpenRecordset StrSql, rs_Datos
    
    If Not rs_Datos.EOF Then
        Monto = IIf(EsNulo(rs_Datos!Monto), 0, rs_Datos!Monto)
        cant = IIf(EsNulo(rs_Datos!cant), 0, rs_Datos!cant)
    End If
    
    rs_Datos.Close
    
    If MontoCant Then
        Valor = Monto
    Else
        Valor = cant
    End If

Set rs_Datos = Nothing
End Sub


Public Sub buscarEstr(ByVal Ternro As Long, ByVal Tenro As Long, ByVal Fecha As Date, ByRef Valor As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca el tipo de estructura de un empleado para una fecha
' Autor      : Martin Ferraro
' Fecha      : 01/10/2010
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Datos As New ADODB.Recordset
Dim aux As String

    aux = ""
    
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " From his_estructura"
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " AND htetdesde <= " & ConvFecha(Fecha)
    StrSql = StrSql & " And (htethasta Is Null Or htethasta >= " & ConvFecha(Fecha) & ")"
    StrSql = StrSql & " And his_estructura.tenro = " & Tenro
    StrSql = StrSql & " And his_estructura.ternro = " & Ternro
    OpenRecordset StrSql, rs_Datos
    
    If Not rs_Datos.EOF Then
        aux = IIf(EsNulo(rs_Datos!estrdabr), "", rs_Datos!estrdabr)
    End If
    
    rs_Datos.Close
    
    Valor = aux

Set rs_Datos = Nothing
End Sub


Public Sub insertarDet(ByVal BproNro As Long, ByVal Ternro As Long, ByVal Valor As String, ByVal Columna As Long, ByVal Orden As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta Registro de detalle
' Autor      : Martin Ferraro
' Fecha      : 01/10/2010
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

    StrSql = "INSERT INTO repgraldet"
    StrSql = StrSql & " (bpronro ,ternro,valor,columna,orden)"
    StrSql = StrSql & " VALUES("
    StrSql = StrSql & "  " & BproNro
    StrSql = StrSql & " ," & Ternro
    StrSql = StrSql & " ,'" & Mid(Valor, 1, 200) & "'"
    StrSql = StrSql & " ," & Columna
    StrSql = StrSql & " ," & Orden
    StrSql = StrSql & " )"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

