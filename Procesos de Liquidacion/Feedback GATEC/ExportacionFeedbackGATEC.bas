Attribute VB_Name = "FeedbackGATEC"
Option Explicit

Global Const Version = "1.00"
Global Const FechaModificacion = "16/06/2014"
Global Const UltimaModificacion = "Version Inicial" 'MDZ - CAS-23486 - TABACAL - Exportación liq GATEC - Custom

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
Global fs, f
'Global Flog
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
Global ConcNro As Integer
Global ConcCod As String
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
' Autor      : MDZ
' Fecha      : 09/06/2014
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim rs As New ADODB.Recordset
Dim PID As String
Dim Parametros As String
Dim ArrParametros

Dim periodo As Long
Dim TipogruLiq As Integer


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
    
    Nombre_Arch = PathFLog & "ExportacionSIGA" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    On Error Resume Next
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Then
        Flog.WriteLine "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Then
        Flog.WriteLine "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0

    On Error GoTo ME_Main
    
    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
   
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.WriteLine "-----------------------------------------------------------------"
    Flog.WriteLine "Version = " & Version
    Flog.WriteLine "Modificacion = " & UltimaModificacion
    Flog.WriteLine "Fecha = " & FechaModificacion
    Flog.WriteLine "-----------------------------------------------------------------"
    Flog.WriteLine
    Flog.WriteLine "PID = " & PID
   
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.WriteLine Espacios(Tabulador * 0) & "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND btprcnro = 420"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
       'Obtengo los parametros del proceso
       Parametros = rs!bprcparam
       ArrParametros = Split(Parametros, "@")
       
       periodo = CLng(ArrParametros(0))
       'TipogruLiq = CInt(ArrParametros(2))
       
       Flog.WriteLine "  Periodo     : " & periodo
       'Flog.writeline "  Opcion      : " & TipogruLiq & " (1.- AGR 2.- PIN)"
       Flog.WriteLine
        
       'Call Generar_Archivo(Periodo, TipogruLiq)
       Call Generar_Exportacion(periodo)
    Else
        Flog.WriteLine Espacios(Tabulador * 0) & "No se encontraron los datos del proceso nro " & NroProceso
    End If
    
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.WriteLine Espacios(Tabulador * 0) & "Proceso Finalizado Correctamente"
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
       Flog.WriteLine Espacios(Tabulador * 0) & "Proceso Finalizado Incompleto"
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords

    TiempoFinalProceso = GetTickCount
    Flog.WriteLine Espacios(Tabulador * 0) & "=================================================="
    Flog.WriteLine Espacios(Tabulador * 0) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.WriteLine Espacios(Tabulador * 0) & "=================================================="
    Flog.Close
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    objconnProgreso.Close
    objConn.Close
Exit Sub
    
ME_Main:
    HuboErrores = True
    Flog.WriteLine "Error: " & Err.Description
    Flog.WriteLine "Ultimo SQL: " & StrSql
End Sub

'Private Sub Generar_Archivo(ByVal pliqnro As Long, ByVal TipogruLiq As Integer)
Private Sub Generar_Exportacion(ByVal pliqnro As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera la exportacion a tabla GATEC_LIQ
' Autor      : MDZ
' Fecha      : 09/06/2014
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim carpeta
Dim fs1

Dim cantRegistros As Long

Dim NroReporte As Integer
'Dim c_fecdesde As String
'Dim c_fechasta As String
Dim lista_estruc As String
Dim lista_AC As String
Dim lista_CO As String
Dim tipo_estructura
Dim v_categ As String
Dim v_cc As String
Dim v_thnro As Long
Dim fechasta As Date
Dim fecdesde As Date
Dim periodo

'campos = ""

Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

    StrSql = " SELECT pliqdesde, pliqhasta, pliqmes, pliqanio FROM periodo WHERE pliqnro=" & pliqnro
    OpenRecordset StrSql, rs
    
    'c_fecdesde = Date
    'c_fechasta = Date
    'fechasta = Date
    If Not rs.EOF Then
        'c_fecdesde = CStr(Format(Year(rs!pliqdesde), "0000") & Format(Month(rs!pliqdesde), "00") & Format(Day(rs!pliqdesde), "00"))
        'c_fechasta = CStr(Format(Year(rs!pliqhasta), "0000") & Format(Month(rs!pliqhasta), "00") & Format(Day(rs!pliqhasta), "00"))
        fechasta = rs!pliqhasta
        fecdesde = rs!pliqdesde
        periodo = Format(rs!pliqmes, "00") & Format(rs!pliqanio, "0000")
    Else
        Flog.WriteLine "ERROR : No se encontró el periodo de liquidacion " & pliqnro
        Exit Sub
    End If
    rs.Close
    
    
    'Configuracion del Reporte
    NroReporte = 439
    StrSql = "SELECT * FROM confrep "
    StrSql = StrSql & " WHERE repnro = " & NroReporte
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    lista_estruc = "0"
    lista_AC = ""
    lista_CO = ""
    tipo_estructura = 32
    
    Dim campos()
    
    ReDim campos(4, 0)
    
    If rs.EOF Then
        Flog.WriteLine "No se encontró la configuración del Reporte"
        Flog.WriteLine "   Se deben configurar 2 tipos de columnas:"
        Flog.WriteLine "     TE : Tipo de estructura   No se encontró la configuración del Reporte. Default 32 (Grupo de Liquidacion). Unico"
        Flog.WriteLine "     EST: Lista de estructuras del tipo anterior. Una o mas"
        Exit Sub
    Else
        Do Until rs.EOF
            Select Case rs!conftipo
            Case "TE":
                tipo_estructura = rs!confval
            Case "EST":
                lista_estruc = lista_estruc & "," & rs!confval
            Case "CO", "CCO":
                Flog.WriteLine UBound(campos, 2)
                ReDim Preserve campos(4, UBound(campos, 2) + 1)
                'es un concepto
                lista_CO = lista_CO & " '" & rs!confval2 & "',"
                'campos = campos & rs!confetiq & ","
                'campos.Add "CO" & rs!confval2, rs!confetiq
                campos(0, UBound(campos, 2) - 1) = "CO"
                campos(1, UBound(campos, 2) - 1) = rs!confval2
                campos(2, UBound(campos, 2) - 1) = rs!confetiq
                If rs!conftipo = "CCO" Then
                    campos(3, UBound(campos, 2) - 1) = 1
                Else
                    campos(3, UBound(campos, 2) - 1) = 0
                End If
                
            Case "AC", "CAC":
                ReDim Preserve campos(4, UBound(campos, 2) + 1)
                'es un acumulador
                lista_AC = lista_AC & " " & rs!confval2 & ","
                'campos = campos & rs!confetiq & ","
                'campos.Add "AC" & rs!confval2, rs!confetiq
                campos(0, UBound(campos, 2) - 1) = "AC"
                campos(1, UBound(campos, 2) - 1) = rs!confval2
                campos(2, UBound(campos, 2) - 1) = rs!confetiq
                
                If rs!conftipo = "CAC" Then
                    campos(3, UBound(campos, 2) - 1) = 1
                Else
                    campos(3, UBound(campos, 2) - 1) = 0
                End If
                
            End Select
            rs.MoveNext
        Loop
    End If
    rs.Close

    If Len(Trim(lista_AC)) > 0 Then
        lista_AC = Left(lista_AC, Len(lista_AC) - 1)
    End If
    
    If Len(Trim(lista_CO)) > 0 Then
        lista_CO = Left(lista_CO, Len(lista_CO) - 1)
    End If
    
    Flog.WriteLine "     "
    
    
    
    StrSql = "SELECT estrcodext, empleado.ternro, empleg, tprocdesc, cliqnro "
    StrSql = StrSql & " FROM proceso "
    StrSql = StrSql & " INNER JOIN tipoproc ON tipoproc.tprocnro = proceso.tprocnro "
    StrSql = StrSql & " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro "
    StrSql = StrSql & " INNER JOIN empleado ON cabliq.empleado = empleado.ternro "
    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro AND "
    StrSql = StrSql & " ((his_estructura.htetdesde <= " & ConvFecha(fecdesde) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(fechasta)
    StrSql = StrSql & " or his_estructura.htethasta >= " & ConvFecha(fecdesde) & ")) OR "
    StrSql = StrSql & " (his_estructura.htetdesde >= " & ConvFecha(fecdesde) & " AND (his_estructura.htetdesde <= " & ConvFecha(fechasta) & ")))"

    'StrSql = StrSql & " AND htetdesde <= " & ConvFecha(fechasta)
    'StrSql = StrSql & " AND (htethasta Is Null Or htethasta <= " & ConvFecha(fecdesde) & ") "
    StrSql = StrSql & " AND his_estructura.tenro = " & tipo_estructura
    StrSql = StrSql & " AND his_estructura.estrnro IN (" & lista_estruc & ")"
    StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
    StrSql = StrSql & " WHERE proceso.pliqnro = " & pliqnro
    StrSql = StrSql & " ORDER BY empleado.empleg "
    'StrSql = StrSql & " ORDER BY proceso.profecini, proceso.pronro, his_estructura.estrnro, empleado.empleg "
    
    'Flog.writeline StrSql
    'Flog.writeline " "
    'Flog.writeline " "
    OpenRecordset StrSql, rs
    
    'seteo de las variables de progreso
    Progreso = 0
    If Not rs.EOF Then
        cantRegistros = rs.RecordCount
        If cantRegistros = 0 Then
           cantRegistros = 1
           Flog.WriteLine Espacios(Tabulador * 1) & "No se encontraron datos a Exportar."
        End If
    Else
        cantRegistros = 1
        Flog.WriteLine Espacios(Tabulador * 1) & "No se encontraron datos a Exportar."
    End If
    IncPorc = (99 / cantRegistros)
    
    Dim i
   
    
    Dim legajo_ant
    legajo_ant = 0
    
    Dim lstCliq
    lstCliq = "0"
    
    Do While Not rs.EOF
        
        'Actualizo el estado del proceso
        TiempoAcumulado = GetTickCount
        cantRegistros = cantRegistros - 1
        Progreso = Progreso + IncPorc
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
       
        
        legajo_ant = rs("empleg")
        lstCliq = lstCliq & "," & rs("cliqnro")
        
        rs.MoveNext
        
        
        If Not rs.EOF Then
            If (legajo_ant <> 0 And rs("empleg") <> legajo_ant) Then
                insertarRegistro lista_CO, lista_AC, lstCliq, legajo_ant, periodo, campos
                lstCliq = "0"
            End If
        Else
            insertarRegistro lista_CO, lista_AC, lstCliq, legajo_ant, periodo, campos
        End If
    Loop
    
    
   
Fin:
    'Cierro y libero todo
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
  
Exit Sub

ME_Local:
    Flog.WriteLine
'    Resume Next
    Flog.WriteLine Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.WriteLine Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.WriteLine Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.WriteLine Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.WriteLine
End Sub


Sub insertarRegistro(lista_CO, lista_AC, lstCliq, Legajo, periodo, campos)
    
    Dim rs2 As New ADODB.Recordset
    Dim sqlCampos
    Dim sqlValores
    sqlValores = ""
    sqlCampos = ""
    
    'leo todos los conceptos en la lista de conceptos y todos los acumuladores en la lista de acumuladores
    ' desde la liquidacion
    If Trim(lista_CO) <> "" Then
        StrSql = "SELECT 'CO' concepto, conccod cocodigo, 0 accodigo, sum(dlimonto) monto, sum(dlicant) cantidad from detliq inner join concepto on "
        StrSql = StrSql & " (detliq.concnro=concepto.concnro AND conccod IN (" & lista_CO & "))"
        StrSql = StrSql & " WHERE detliq.cliqnro IN (" & lstCliq & ") GROUP BY conccod"
    End If
    
    If Trim(lista_CO) <> "" And Trim(lista_AC) <> "" Then
        StrSql = StrSql & " UNION "
    End If
    
    If Trim(lista_AC) <> "" Then
        StrSql = StrSql & " select 'AC' concepto, '0' cocodigo, acunro accodigo, sum(almonto) monto, sum(alcant) cantidad from acu_liq WHERE acunro IN (" & lista_AC & ")"
        StrSql = StrSql & " AND cliqnro IN  (" & lstCliq & ") GROUP BY acunro"
    End If
    
    'Flog.writeline StrSql
    OpenRecordset StrSql, rs2
    
    'genero el insert a la tabla GATEC_LIQ (si el registro existe lo actualizo)
    sqlCampos = "COD_EMPR, FPG_FUNC, FPG_MES, "
    sqlValores = "'01', " & Legajo & ", '" & periodo & "',"
        
    Dim i
    Dim inserto As Boolean
    inserto = False
    'Dim arrCaampos
    'arrCaampos = Split(campos)
    Do Until rs2.EOF
        'Flog.writeline rs2!concepto & "  AC:" & rs2!accodigo & "  CO:" & rs2!cocodigo & "  monto:" & rs2!Monto
        
        For i = 0 To UBound(campos, 2) - 1
            'Flog.writeline campos(2, i)
            
           
            If rs2!concepto = "CO" Then
                If campos(0, i) = "CO" And _
                   campos(1, i) = rs2!cocodigo Then
                     
                    'Flog.writeline campos(2, i) & "=" & rs2!Monto
                    
                    sqlCampos = sqlCampos & campos(2, i) & ","
                    If campos(3, i) = 0 Then
                        sqlValores = sqlValores & rs2!Monto & ","
                    Else
                        sqlValores = sqlValores & rs2!Cantidad & ","
                    End If
                    inserto = True
                    
                End If
            
            Else
                If campos(0, i) = "AC" And _
                   CInt(campos(1, i)) = CInt(rs2!accodigo) Then
                    
                    'Flog.writeline campos(2, i) & "=" & rs2!Monto
                    
                    sqlCampos = sqlCampos & campos(2, i) & ","
                    
                    If campos(3, i) = 0 Then
                        sqlValores = sqlValores & rs2!Monto & ","
                    Else
                        sqlValores = sqlValores & rs2!Cantidad & ","
                    End If
                    
                    inserto = True
                    
                End If
            
            End If
            
        Next
        
        rs2.MoveNext
    Loop
    
    'elimino registro anterior
    StrSql = "DELETE FROM GATEC_LIQ WHERE COD_EMPR='01' AND FPG_FUNC=" & Legajo & " AND FPG_MES=" & periodo
    objConn.Execute StrSql, , adExecuteNoRecords

    If inserto Then
    
        'saco las ultimas ","
        sqlCampos = Left(sqlCampos, Len(sqlCampos) - 1)
        sqlValores = Left(sqlValores, Len(sqlValores) - 1)
        
        StrSql = "INSERT INTO  GATEC_LIQ  (" & sqlCampos & ") VALUES (" & sqlValores & ")"
        'Flog.writeline StrSql
        
        objConn.Execute StrSql, , adExecuteNoRecords
    
    End If
    
End Sub

Sub imprimirTexto(Texto, archivo, Longitud, derecha)
'Rutina genérica para imprimir un TEXTO, de una LONGITUD determinada.
'Los sobrantes se rellenan con CARACTER

Dim cadena
Dim txt
Dim u
Dim longTexto
    
    If IsNull(Texto) Then
        longTexto = 1
        cadena = " "
    Else
        longTexto = Len(Texto)
        cadena = Mid(CStr(Texto), 1, Longitud) & String(Longitud - longTexto, " ")
    End If
    
    archivo.Write cadena
    
End Sub


