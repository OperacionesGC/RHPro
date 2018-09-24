Attribute VB_Name = "ExpUocra"
' ---------------------------------------------------------------------------------------------
' Descripcion: Proyecto encargado de generar la exportacion Interface Uocra
' Autor      : Lisandro Moro
' Fecha      : 07/11/2007
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "07/11/2007"
'Global Const UltimaModificacion = " " 'Version Inicial

'Global Const Version = "1.01"
'Global Const FechaModificacion = "05/12/2007"
'Global Const UltimaModificacion = " " 'Buscaba mal el cuil

'Global Const Version = "1.02"
'Global Const FechaModificacion = "31/07/2009"
'Global Const UltimaModificacion = " " 'MB - Encriptacion de string connection


'Global Const Version = "1.03"
'Global Const FechaModificacion = "20/02/2015"
'Global Const UltimaModificacion = " " 'MDF - CAS-29520 - KOMPENDIUM - Error en reporte UOCRA
                                      ' control sobre valores nulos
                                      

'Global Const Version = "1.04"
'Global Const FechaModificacion = "02/03/2015"
'Global Const UltimaModificacion = " " 'MDF - CAS-29520 - KOMPENDIUM - Error en reporte UOCRA
                                      'se busca por codigo de concepto
                                      
'Global Const Version = "1.05"
'Global Const FechaModificacion = "15/04/2015"
'Global Const UltimaModificacion = " " 'FB - CAS-29963 - HOLDEC - BUG Exportacion UOCRA
                                      'Se modifica la consulta para contemplar sólo los empleados liquidados en los procesos seleccionados

Global Const Version = "1.06"
Global Const FechaModificacion = "29/06/2015"
Global Const UltimaModificacion = " " 'Miriam Ruiz - CAS-29520 - KOMPENDIUM - Cambio en reporte UOCRA
                                      'Se agrega como último campo de la exportación si la empresa pertenece o no a la administración pública

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

'Global tenro1 As Integer
'Global estrnro1 As Integer
'Global tenro2 As Integer
'Global estrnro2 As Integer
'Global tenro3 As Integer
'Global estrnro3 As Integer

Global errorConfrep As Boolean

Global TipoCols4(200)
Global CodCols4(200)
Global TipoCols5(200)
Global CodCols5(200)

Global Mes1 As String
Global mesPorc1 As String
Global Mes2 As String
Global mesPorc2 As String
Global Mes3 As String
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
' Autor      : LIsandro Moro
' Fecha      : 05/11/2007
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim PID As String
Dim Parametros As String
Dim ArrParametros

'Dim Sucursal As Long
'Dim Periodo As Long



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

    'Flog.writeline "licho"
 
    On Error Resume Next
    'Abro la conexion
    Flog.writeline strconexion
    
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
    Flog.writeline "err:" & Err.Description
    On Error GoTo 0

    On Error GoTo ME_Main
    
    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ExportacionUocra" & "-" & NroProceso & ".log"
    
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
   
    '-----------------------------------------------mdf
    
    
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

    
    
    
    
    
    
    '-----------------------------------------------mdf
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 0) & "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND btprcnro = " & 206
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
       'Obtengo los parametros del proceso
        Parametros = rs!bprcparam
        
        If Not IsNull(Parametros) Then
            ArrParametros = Split(Parametros, "@")
       
            'Sucursal = CLng(ArrParametros(0))
            'Periodo = CLng(ArrParametros(1))
      
            Call Generar_Archivo(ArrParametros)
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

Function cambiarFormato(cadena, separador)

 Dim l_salida
 
 l_salida = cadena

 If (InStr(cadena, ".")) Then
    If separador = "," Then
       l_salida = Replace(cadena, ".", ",")
    End If
 Else
     If (InStr(cadena, ",")) Then
        If separador = "." Then
           l_salida = Replace(cadena, ",", ".")
        End If
     End If
 End If

 cambiarFormato = l_salida

End Function 'cambiarFormato(cadena,separador)

Sub imprimirTexto(Texto, archivo, Longitud, derecha)
'Rutina genérica para imprimir un TEXTO, de una LONGITUD determinada.
'Los sobrantes se rellenan con CARACTER

Dim cadena
Dim txt
Dim u
Dim longTexto
    
    If IsNull(Texto) Then
        longTexto = 0
        cadena = " "
    Else
        longTexto = Len(Texto)
        cadena = CStr(Texto)
    End If
    
    u = CInt(Longitud) - longTexto
    If u < 0 Then
        archivo.Write cadena
    Else
        If derecha Then
             archivo.Write cadena & String(u, " ")
        Else
            archivo.Write String(u, " ") & cadena
        End If
    End If

End Sub

Sub imprimirTextoNro(ByVal Texto, ByRef archivo, ByVal Longitud, ByVal derecha As Boolean)
'Rutina genérica para imprimir un TEXTO, de una LONGITUD determinada.
'Los sobrantes se rellenan con CARACTER

Dim cadena
Dim txt
Dim u
Dim longTexto
    
    If IsNull(Texto) Then
        longTexto = 0
        cadena = " "
    Else
        longTexto = Len(Texto)
        cadena = CStr(Texto)
    End If
    
    u = CInt(Longitud) - longTexto
    If u < 0 Then
        archivo.Write cadena
    Else
        If derecha Then
            archivo.Write cadena & String(u, " ")
        Else
            archivo.Write String(u, " ") & cadena
        End If
    End If

End Sub

Sub imprimirTextoConCeros(Texto, archivo, Longitud, derecha)
'Rutina genérica para imprimir un TEXTO, de una LONGITUD determinada.
'Los sobrantes se rellenan con CARACTER

Dim cadena
Dim txt
Dim u
Dim longTexto
    
    If IsNull(Texto) Then
        longTexto = 1
        cadena = "0"
    Else
        longTexto = Len(Texto)
        cadena = CStr(Texto)
    End If
    
    archivo.Write cadena
    
End Sub

Function CerozIzq(Valor)
Dim aux
Dim result
    
    If InStr(Valor, ".") = 0 Then
        result = Valor & ".00"
    Else
        aux = Mid(Valor, InStr(Valor, "."), Len(Valor))
        If Len(aux) = 2 Then
            result = Valor & "0"
        Else
            result = Valor
        End If
    End If
        
    CerozIzq = result
    
End Function

Private Sub Generar_Archivo(ByVal ArrParametros As Variant)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera la exportacion
' Autor      : Lisandro Moro
' Fecha      : 07/11/2007
' Ultima Mod :
' Descripcion:

'CUIT                       11
'AFILIADO                   1
'CUOTA_SINDICAL             8
'FONDO_CESE_LABORAL         8
'FECHA_INGRESO              8
'CP_LABORAL                 4
'CONVENIO                   2
'CATEGORIA                  2

' ---------------------------------------------------------------------------------------------
'Dim objRs As New ADODB.Recordset
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'variables recibidas de parametro
Dim l_filtro As String
Dim l_fecestr As Date
Dim l_tenro1 As Integer
Dim l_estrnro1 As Integer
Dim l_tenro2 As Integer
Dim l_estrnro2 As Integer
Dim l_tenro3 As Integer
Dim l_estrnro3 As Integer

Dim l_pliqdesde As Long
Dim l_pliqhasta As Long
Dim l_desde As Date
Dim l_hasta As Date
Dim l_proaprob As Integer
Dim l_listaproc As String
Dim l_concnro As Long
Dim l_empresa As Long
Dim l_conceptonombre As String
Dim l_orden As String
Dim l_lista_emp As String
Dim StrSql2 As String
'Dim filtro As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Nombre_Arch As String
Dim NroModelo As Long
Dim directorio As String
Dim carpeta
Dim fs1

Dim Sep As String
Dim SepDec As String

Dim cantRegistros As Long
Dim Empresa As String

Dim Legajo As String
Dim Tercero As Long
Dim estrUOCRA As Long

Dim cuit As String '11 caracteres sin guiones. Ej. 20244367918
Dim AFILIADO As String ' SI(CUOTA SINDICAL <> 0;"S";"N") . Identifica si esta afiliado o no.
Dim CUOTA_SINDICAL As String ' ? de concepto y acum. (config en confrep)
Dim FONDO_CESE_LABORAL As String ' ? de concepto y acum. (config en confrep)
Dim FECHA_INGRESO As String 'fecha de ingreso de empleado en formato ddmmyyyy. Fase con fecha de alta reconocida. Si no hay ninguna fase con la marca Ò Empleado.empfecaltgr.
Dim CP_LABORAL As String ' Código postal de la sucursal del empleado. Podemos asociarle un tipo de código X a las sucursales y buscar y mostrar ese código. Revisar Reporte fondo de desempleo porque creo que ya existe el tipo de codigo.
Dim CONVENIO As String ' Podemos asociarle un tipo de código X a los convenios y escribir ese código. Revisar Reporte fondo de desempleo porque creo que ya existe el tipo de codigo.
Dim CATEGORIA As String ' Podemos asociarle un tipo de código X a las categorías y escribir ese código. Revisar Reporte fondo de desempleo porque creo que ya existe el tipo de codigo.

Dim NroReporte As Long
Dim Acumulador As Long

Dim l_almonto As Double
Dim l_dlimonto As Double
Dim l_almonto_cuota As Double
Dim l_dlimonto_fondo As Double
Dim l_almonto_fondo As Double
Dim l_dlimonto_cuota As Double
Dim l_Suma_Cuota As Double
Dim l_Suma_Fondo As Double
Dim Antiguedad As Long

Dim l_Lista_ternro As String
Dim l_Lista_ternro_temp
Dim I As Long

Dim Lista_Sindicatos As String
Dim Lista_conceptos_Fondo As String
Dim Lista_Acumuladores_Fondo As String
Dim Lista_conceptos_Cuota As String
Dim Lista_Acumuladores_Cuota As String

Dim arrSindicatos
Dim formatofecha As String
formatofecha = "DDMMYYYY"

Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs_Aux As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim Admpublica As String

    
    On Error GoTo ME_Local
    
    '////////////////////////////////////////////////////////////////////
        'filtro empleg - (0 = 0) AND (empleg >= 1) AND (empleg <= 9999999999)
        l_filtro = CStr(ArrParametros(0))
        'StrSql2 = l_filtro
        'fecha estructura
        l_fecestr = CDate(ArrParametros(1))
        'Estructura 1
        If ArrParametros(2) <> "" Then
         l_tenro1 = CInt(ArrParametros(2))
        Else
         l_tenro1 = 0
        End If
        If ArrParametros(3) <> "" Then
         l_estrnro1 = CInt(ArrParametros(3))
        Else
          l_estrnro1 = 0
        End If
        'Estructura 2
        If ArrParametros(4) <> "" Then
         l_tenro2 = CInt(ArrParametros(4))
        Else
         l_tenro2 = 0
        End If
        If ArrParametros(5) <> "" Then
         l_estrnro2 = CInt(ArrParametros(5))
        Else
          l_estrnro2 = 0
        End If
        'Estructura 3
        If ArrParametros(6) <> "" Then
         l_tenro3 = CInt(ArrParametros(6))
        Else
         l_tenro3 = 0
        End If
        If ArrParametros(7) <> "" Then
          l_estrnro3 = CInt(ArrParametros(7))
        Else
          l_estrnro3 = 0
        End If
        'Periodo liquidacion desde
        l_pliqdesde = CLng(IIf(ArrParametros(8) = "", 0, ArrParametros(8)))
        'Periodo liquidacion hasta
        l_pliqhasta = CLng(IIf(ArrParametros(9) = "", 0, ArrParametros(9)))
        'fecha periodo desde
        l_desde = CDate(ArrParametros(10))
        'fecha periodo hasta
        l_hasta = CDate(ArrParametros(11))
        'Procesos estados
        '   Aprob.Definitivo = 3
        '   Aprob.Provisorio = 2
        '   Liquidado = 1
        '   No Liquidado = 0
        '   Todos = -1
        l_proaprob = CLng(IIf(ArrParametros(12) = "", 0, ArrParametros(12))) 'CInt(ArrParametros(12))
        ' Lista de Procesos
        l_listaproc = CStr(ArrParametros(13))
        ' Conseptos - no se usa
        'l_concnro = CLng(ArrParametros(14))
        'Empresa nro
        l_empresa = CLng(ArrParametros(15))
        'Nombre concepto no se usa
        'l_conceptonombre = CStr(ArrParametros(16))
        'Orden
        l_orden = CStr(ArrParametros(17))
        'Lista de empleados
        l_lista_emp = CStr(ArrParametros(18))
        
        l_Lista_ternro = "0"
        If l_lista_emp <> "0" Then
            l_Lista_ternro_temp = Split(l_lista_emp, ",")
            For I = 0 To UBound(l_Lista_ternro_temp) - 1
                l_Lista_ternro = l_Lista_ternro & "," & l_Lista_ternro_temp(I + 1)
                I = I + 1
            Next I
        End If
    '////////////////////////////////////////////////////////////////////
    
    
    NroModelo = 917 'lichok xxx
    'Directorio de exportacion
    StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        directorio = Trim(rs!sis_dirsalidas)
    End If

    StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
    OpenRecordset StrSql, rs_Modelo
    If Not rs_Modelo.EOF Then
        If Not IsNull(rs_Modelo!modarchdefault) Then
            directorio = directorio & Trim(rs_Modelo!modarchdefault)
        Else
            Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta desteino. El archivo será generado en el directorio default"
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo " & NroModelo & ". El archivo será generado en el directorio default"
    End If

    'Obtengo los datos del separador
    Sep = rs_Modelo!modseparador
    SepDec = rs_Modelo!modsepdec
    Flog.writeline Espacios(Tabulador * 0) & "Separador seleccionado: " & Sep

    Nombre_Arch = directorio & "\ExpUocra" & "-" & NroProceso & ".txt"
    Flog.writeline Espacios(Tabulador * 1) & "Se crea el archivo: " & Nombre_Arch
    Set fs = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
        Set carpeta = fs1.CreateFolder(directorio)
    End If
    'desactivo el manejador de errores
    Set ArchExp = fs.CreateTextFile(Nombre_Arch, True)


    On Error GoTo ME_Local

    'Configuracion del Reporte
    NroReporte = 220
    
    Lista_Sindicatos = "0"
    Lista_Acumuladores_Cuota = "0"
    Lista_conceptos_Fondo = "0"
    Lista_Acumuladores_Fondo = "0"
    Lista_conceptos_Cuota = "0"
    
    StrSql = "SELECT * FROM confrep "
    StrSql = StrSql & " WHERE repnro = " & NroReporte
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Flog.writeline " No hay Conceptos ni Acumuladores configurados "
        'Exit Sub
    Else
        Do While Not rs.EOF
            Select Case UCase(rs("conftipo"))
                Case "EST"
                    Lista_Sindicatos = Lista_Sindicatos & "," & rs!confval
                Case "CO"
                    If UCase(Left(CStr(rs("confetiq")), 5)) = "CUOTA" Then 'CUOTA
                        If IsNull(rs!confval2) Or rs!confval2 = "" Then
                            Lista_conceptos_Cuota = Lista_conceptos_Cuota & "," & rs!confval
                        Else
                            Lista_conceptos_Cuota = Lista_conceptos_Cuota & "," & rs!confval2
                        End If
                    ElseIf UCase(Left(CStr(rs("confetiq")), 5)) = "FONDO" Then 'FONDO
                        If IsNull(rs!confval2) Or rs!confval2 = "" Then
                            Lista_conceptos_Fondo = Lista_conceptos_Fondo & "," & rs!confval
                        Else
                            Lista_conceptos_Fondo = Lista_conceptos_Fondo & "," & rs!confval2
                        End If
                    Else
                        Flog.writeline " Mal configurado el concepto " & rs("confetiq")
                    End If
                Case "AC"
                    If UCase(Left(CStr(rs("confetiq")), 5)) = "CUOTA" Then 'CUOTA
                        Lista_Acumuladores_Cuota = Lista_Acumuladores_Cuota & "," & rs!confval
                    ElseIf UCase(Left(CStr(rs("confetiq")), 5)) = "FONDO" Then 'FONDO
                        Lista_Acumuladores_Fondo = Lista_Acumuladores_Fondo & "," & rs!confval
                    Else
                        Flog.writeline " Mal configurado el Acumulador " & rs("confetiq")
                    End If
            End Select
            rs.MoveNext
        Loop
    End If
    
    'Busco si la empresa es de la administración pública
    
    StrSql = "SELECT * FROM empresa "
    StrSql = StrSql & " INNER JOIN tipempdor ON empresa.tipempnro = tipempdor.tipempnro "
    StrSql = StrSql & " AND tipempdesabr = 'Administración Pública' "
    StrSql = StrSql & " AND empresa.estrnro = " & l_empresa
     Flog.writeline " Administracion publica: " & StrSql
    OpenRecordset StrSql, rs2
    If rs2.EOF Then
         Admpublica = "N"
    Else
        Admpublica = "S"
    End If
    
    
    If l_tenro3 <> 0 Then ' esto ocurre solo cuando se seleccionan los tres niveles
        StrSql = " SELECT DISTINCT empleado.ternro "
        StrSql = StrSql & ", estact1.tenro tenro1, estact1.estrnro estrnro1 "
        StrSql = StrSql & ", estact2.tenro tenro2, estact2.estrnro estrnro2 "
        StrSql = StrSql & ", estact3.tenro tenro3, estact3.estrnro estrnro3 "
        StrSql = StrSql & " FROM empleado "
        'FB
        StrSql = StrSql & " INNER JOIN cabliq ON empleado.ternro = cabliq.empleado "
        'FB
        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro AND estact1.tenro  = " & l_tenro1
        StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(l_fecestr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(l_fecestr) & "))"
        If l_estrnro1 <> 0 Then 'cuando se le asigna un valor al nivel 1
                StrSql = StrSql & " AND estact1.estrnro = " & l_estrnro1
        End If
        StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & l_tenro2
        StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(l_fecestr) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(l_fecestr) & "))"
        If l_estrnro2 <> 0 Then 'cuando se le asigna un valor al nivel 2
            StrSql = StrSql & " AND estact2.estrnro =" & l_estrnro2
        End If
        StrSql = StrSql & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro  = " & l_tenro3
        StrSql = StrSql & " AND (estact3.htetdesde<=" & ConvFecha(l_fecestr) & " AND (estact3.htethasta is null or estact3.htethasta>=" & ConvFecha(l_fecestr) & "))"
        If l_estrnro3 <> 0 Then ' cuando se le asigna un valor al nivel 3
                StrSql = StrSql & " AND estact3.estrnro =" & l_estrnro3
        End If
        StrSql = StrSql & " WHERE " & l_filtro
        'If Len(StrSql2) <> 0 Then
        '    StrSql = StrSql & " AND (" & StrSql2 & ") "
        'End If
        If l_Lista_ternro <> "0" Then
            StrSql = StrSql & " AND empleado.ternro IN (" & l_Lista_ternro & ")"
        End If
        'FB - 14/04/2015
        If l_listaproc <> "0,0" Then
            StrSql = StrSql & " AND cabliq.pronro IN (" & l_listaproc & ")"
        End If
        'FB
        'StrSql = StrSql & " ORDER BY tenro1,estrnro1,tenro2,estrnro2,tenro3,estrnro3," & l_orden & ", empleado.ternro"
    ElseIf l_tenro2 <> 0 Then  'ocurre cuando se selecciono hasta el segundo nivel
        StrSql = "SELECT DISTINCT empleado.ternro, empleg "
        StrSql = StrSql & ", estact1.tenro tenro1, estact1.estrnro estrnro1 "
        StrSql = StrSql & ", estact2.tenro tenro2, estact2.estrnro estrnro2 "
        StrSql = StrSql & " FROM empleado  "
        'FB
        StrSql = StrSql & " INNER JOIN cabliq ON empleado.ternro = cabliq.empleado "
        'FB
        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro AND estact1.tenro = " & l_tenro1
        StrSql = StrSql & " AND (estact1.htetdesde <= " & ConvFecha(l_fecestr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(l_fecestr) & "))"
        If l_estrnro1 <> 0 Then
            StrSql = StrSql & " AND estact1.estrnro =" & l_estrnro1
        End If
        StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & l_tenro2
        StrSql = StrSql & " AND (estact2.htetdesde <= " & ConvFecha(l_fecestr) & " AND (estact2.htethasta is null or estact2.htethasta >= " & ConvFecha(l_fecestr) & "))"
        If l_estrnro2 <> 0 Then
            StrSql = StrSql & " AND estact2.estrnro = " & l_estrnro2
        End If
        StrSql = StrSql & " WHERE " & l_filtro
        'If Len(StrSql2) <> 0 Then
        '    StrSql = StrSql & " AND (" & StrSql2 & ") "
        'End If
        If l_Lista_ternro <> "0" Then
            StrSql = StrSql & " AND empleado.ternro IN (" & l_Lista_ternro & ")"
        End If
        'FB - 14/04/2015
        If l_listaproc <> "0,0" Then
            StrSql = StrSql & " AND cabliq.pronro IN (" & l_listaproc & ")"
        End If
        'FB
        'StrSql = StrSql & " ORDER BY tenro1,estrnro1,tenro2,estrnro2," & l_orden & ", empleado.ternro"
           
    ElseIf l_tenro1 <> 0 Then  ' Cuando solo selecionamos el primer nivel
        StrSql = "SELECT DISTINCT empleado.ternro, empleg "
        StrSql = StrSql & ", estact1.tenro tenro1, estact1.estrnro estrnro1 "
        StrSql = StrSql & " FROM empleado  "
        'FB
        StrSql = StrSql & " INNER JOIN cabliq ON empleado.ternro = cabliq.empleado "
        'FB
        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro AND estact1.tenro = " & l_tenro1
        StrSql = StrSql & " AND (estact1.htetdesde <= " & ConvFecha(l_fecestr) & " AND (estact1.htethasta is null or estact1.htethasta >= " & ConvFecha(l_fecestr) & "))"
        If l_estrnro1 <> 0 Then
            StrSql = StrSql & " AND estact1.estrnro = " & l_estrnro1
        End If
        StrSql = StrSql & " WHERE " & l_filtro
        'If Len(StrSql2) <> 0 Then
        '    StrSql = StrSql & " AND (" & StrSql2 & ") "
        'End If
        If l_Lista_ternro <> "0" Then
            StrSql = StrSql & " AND empleado.ternro IN (" & l_Lista_ternro & ")"
        End If
        'FB - 14/04/2015
        If l_listaproc <> "0,0" Then
            StrSql = StrSql & " AND cabliq.pronro IN (" & l_listaproc & ")"
        End If
        'FB
        'StrSql = StrSql & " ORDER BY tenro1,estrnro1," & l_orden & ", empleado.ternro"
        
    Else  ' cuando no hay nivel de estructura seleccionado
        StrSql = " SELECT DISTINCT empleado.ternro, empleg "
        StrSql = StrSql & " FROM empleado "
        'FB
        StrSql = StrSql & " INNER JOIN cabliq ON empleado.ternro = cabliq.empleado "
        'FB
        StrSql = StrSql & " WHERE  " & l_filtro
        'If Len(StrSql2) <> 0 Then
        '    StrSql = StrSql & " AND (" & StrSql2 & ") "
        'End If
        If l_Lista_ternro <> "0" Then
            StrSql = StrSql & " AND empleado.ternro IN (" & l_Lista_ternro & ")"
        End If
        'FB - 14/04/2015
        If l_listaproc <> "0,0" Then
            StrSql = StrSql & " AND cabliq.pronro IN (" & l_listaproc & ")"
        End If
        'FB
        'StrSql = StrSql & " ORDER BY " & l_orden & ", empleado.ternro"
    End If
                      

    
    'Busco los empleados
    OpenRecordset StrSql, rs
    
    ' _________________________________________________________________________
    Flog.writeline "  SQL para control de los empleados periodo de las cuotas. "
    Flog.writeline "    " & StrSql
    Flog.writeline " "

    If rs.EOF Then
        cantRegistros = 1
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos a Exportar."
        
        Flog.writeline "No se encontraron Empleados para el Reporte."
        StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
        Exit Sub
    Else
        
        cantRegistros = rs.RecordCount
        If cantRegistros = 0 Then
           cantRegistros = 1
           Flog.writeline Espacios(Tabulador * 1) & "Se encontraron " & cantRegistros & " datos a Exportar."
        End If
        IncPorc = (99 / cantRegistros)
            
            Do Until rs.EOF
                
                
                l_dlimonto = 0
                '-Conceptos cuotas---------------------------------------------
                l_dlimonto_cuota = 0
                StrSql = " SELECT dlimonto "
                StrSql = StrSql & " FROM empleado"
                StrSql = StrSql & " INNER JOIN cabliq ON empleado.ternro = cabliq.empleado "
                StrSql = StrSql & " INNER JOIN detliq  ON detliq.cliqnro  = cabliq.cliqnro   "
                StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro  = cabliq.pronro "
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " WHERE concepto.conccod IN (" & Lista_conceptos_Cuota & ")"
                If CStr(l_desde) <> "" Then
                    StrSql = StrSql & " AND periodo.pliqdesde  >= " & ConvFecha(l_desde)
                End If
                If CStr(l_hasta) <> "" Then
                    StrSql = StrSql & " AND periodo.pliqhasta  <= " & ConvFecha(l_hasta)
                End If
                If l_listaproc <> "0,0" Then
                    StrSql = StrSql & " AND cabliq.pronro IN (" & l_listaproc & ")"
                End If
                StrSql = StrSql & " AND empleado.ternro = " & rs("ternro")
                Flog.writeline "sql---->" & StrSql
                OpenRecordset StrSql, rs2
                If rs2.EOF Then
                    Flog.writeline "No se encontraron conceptos de cuotas para el empleado.ternro(" & rs("ternro") & ")"
                Else
                    Do While Not rs2.EOF
                        If Not IsNull(rs2("dlimonto")) Then
                            l_dlimonto_cuota = l_dlimonto_cuota + CDbl(rs2("dlimonto"))
                        End If
                        rs2.MoveNext
                    Loop
                End If
                rs2.Close
                
                '-Conceptos fondo----------------------------------------------
                l_dlimonto_fondo = 0
                StrSql = " SELECT dlimonto "
                StrSql = StrSql & " FROM empleado"
                StrSql = StrSql & " INNER JOIN cabliq ON empleado.ternro = cabliq.empleado "
                StrSql = StrSql & " INNER JOIN detliq  ON detliq.cliqnro  = cabliq.cliqnro   "
                StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro  = cabliq.pronro "
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " WHERE concepto.conccod IN (" & Lista_conceptos_Fondo & ")"
                If CStr(l_desde) <> "" Then
                    StrSql = StrSql & " AND periodo.pliqdesde  >= " & ConvFecha(l_desde)
                End If
                If CStr(l_hasta) <> "" Then
                    StrSql = StrSql & " AND periodo.pliqhasta  <= " & ConvFecha(l_hasta)
                End If
                If l_listaproc <> "0,0" Then
                    StrSql = StrSql & " AND cabliq.pronro IN (" & l_listaproc & ")"
                End If
                StrSql = StrSql & " AND empleado.ternro = " & rs("ternro")
                Flog.writeline "sql---->" & StrSql
                OpenRecordset StrSql, rs2
                If rs2.EOF Then
                    Flog.writeline "No se encontraron conceptos de fondos para el empleado.ternro(" & rs("ternro") & ")"
                Else
                    Do While Not rs2.EOF
                        If Not IsNull(rs2("dlimonto")) Then
                            l_dlimonto_fondo = l_dlimonto_fondo + CDbl(rs2("dlimonto"))
                        End If
                        rs2.MoveNext
                    Loop
                End If
                rs2.Close
                
                
                '**************************************************************
                
                l_almonto = 0
                'Acumuladores cuotas-------------------------------------------
                l_almonto_cuota = 0
                StrSql = " SELECT almonto"
                StrSql = StrSql & " FROM empleado"
                StrSql = StrSql & " INNER JOIN cabliq ON empleado.ternro = cabliq.empleado "
                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro  = cabliq.cliqnro "
                StrSql = StrSql & " INNER JOIN acumulador ON acumulador.acunro = acu_liq.acunro "
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro  = cabliq.pronro "
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " WHERE acu_liq.acunro IN (" & Lista_Acumuladores_Cuota & ")"
                If CStr(l_desde) <> "" Then
                    StrSql = StrSql & " AND periodo.pliqdesde  >= " & ConvFecha(l_desde)
                End If
                If CStr(l_hasta) <> "" Then
                    StrSql = StrSql & " AND periodo.pliqhasta  <= " & ConvFecha(l_hasta)
                End If
                If l_listaproc <> "0,0" Then
                    StrSql = StrSql & " AND cabliq.pronro IN (" & l_listaproc & ")"
                End If
                StrSql = StrSql & " AND empleado.ternro = " & rs("ternro")
                OpenRecordset StrSql, rs2
                If rs2.EOF Then
                    Flog.writeline "No se encontraron Acumuladores cuotas para el empleado.ternro(" & rs("ternro") & ")"
                Else
                    Do While Not rs2.EOF
                        If Not IsNull(rs2("almonto")) Then
                            l_almonto_cuota = l_almonto_cuota + CDbl(rs2("almonto"))
                        End If
                        rs2.MoveNext
                    Loop
                End If
                rs2.Close


                'Acumuladores Fondos-------------------------------------------
                l_almonto_fondo = 0
                StrSql = " SELECT almonto"
                StrSql = StrSql & " FROM empleado"
                StrSql = StrSql & " INNER JOIN cabliq ON empleado.ternro = cabliq.empleado "
                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro  = cabliq.cliqnro "
                StrSql = StrSql & " INNER JOIN acumulador ON acumulador.acunro = acu_liq.acunro "
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro  = cabliq.pronro "
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " WHERE acu_liq.acunro IN (" & Lista_Acumuladores_Fondo & ")"
                If CStr(l_desde) <> "" Then
                    StrSql = StrSql & " AND periodo.pliqdesde  >= " & ConvFecha(l_desde)
                End If
                If CStr(l_hasta) <> "" Then
                    StrSql = StrSql & " AND periodo.pliqhasta  <= " & ConvFecha(l_hasta)
                End If
                If l_listaproc <> "0,0" Then
                    StrSql = StrSql & " AND cabliq.pronro IN (" & l_listaproc & ")"
                End If
                StrSql = StrSql & " AND empleado.ternro = " & rs("ternro")
                OpenRecordset StrSql, rs2
                If rs2.EOF Then
                    Flog.writeline "No se encontraron Acumuladores fondos para el empleado.ternro(" & rs("ternro") & ")"
                Else
                    Do While Not rs2.EOF
                        If Not IsNull(rs2("almonto")) Then
                            l_almonto_fondo = l_almonto_fondo + CDbl(rs2("almonto"))
                        End If
                        rs2.MoveNext
                    Loop
                End If
                rs2.Close

                ' Sumatorias de Conceptos y Acumulados para cuotas y fondos ----------
                l_Suma_Cuota = l_dlimonto_cuota + l_almonto_cuota
                l_Suma_Fondo = l_dlimonto_fondo + l_almonto_fondo
                'If Antiguedad > 356 Then '8%
                'Else '12%
                'End If
                CUOTA_SINDICAL = Right("00000000" & CStr(Round(l_Suma_Cuota, 2) * 100), 8)
                FONDO_CESE_LABORAL = Right("00000000" & CStr(Round(l_Suma_Fondo, 2) * 100), 8)
                
                
                '---------------------------------------------------------------------
                
                'Busco el cuit del empleado
                StrSql = " SELECT nrodoc "
                StrSql = StrSql & " FROM ter_doc "
                StrSql = StrSql & " WHERE ternro = " & rs("ternro")
                StrSql = StrSql & " AND tidnro = 10 " 'aca habia un 6
                OpenRecordset StrSql, rs2
                If rs2.EOF Then
                    Flog.writeline "No se encontraron el cuit para el empleado.ternro(" & rs("ternro") & ")"
                    cuit = "00000000000"
                Else
                    cuit = Left(CStr(rs2("nrodoc")), 13)
                    cuit = Replace(CStr(cuit), "-", "")
                    cuit = Right("00000000" & cuit, 11)
                    cuit = Left(CStr(cuit), 11)
                End If
                rs2.Close
                
                'Lo busco Si pertenece al sindicato UOCRA
                AFILIADO = "N"
                StrSql = " SELECT estrnro "
                StrSql = StrSql & " FROM empleado  "
                StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.tenro = 16 " 'SINDICATO
                StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(l_fecestr) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(l_fecestr) & "))"
                StrSql = StrSql & " WHERE empleado.ternro = " & rs("ternro")
                'StrSql = StrSql & " AND estrnro = " & estrUOCRA
                OpenRecordset StrSql, rs2
                If rs2.EOF Then
                    AFILIADO = "N"
                    Flog.writeline "No se encontraron codigo uocra de la estructura Sindicato para el empleado.ternro(" & rs("ternro") & ")"
                Else
                    arrSindicatos = Split(Lista_Sindicatos, ",")
                    For I = 0 To UBound(arrSindicatos)
                        If CLng(rs2("estrnro")) = CLng(arrSindicatos(I)) Then
                            AFILIADO = "S"
                        End If
                    Next
                End If
                
                
                'Busco la fecha de alta
                Antiguedad = 0
                FECHA_INGRESO = "00000000"
                StrSql = " SELECT altfec "
                StrSql = StrSql & " FROM fases "
                StrSql = StrSql & " WHERE empleado = " & rs("ternro")
                StrSql = StrSql & " AND fasrecofec = -1 "
                OpenRecordset StrSql, rs2
                If rs2.EOF Then
                    rs2.Close
                    Flog.writeline "No se encontraron la fase con fecha de alta reconocida para el empleado.ternro(" & rs("ternro") & ")"
                    StrSql = " SELECT empfaltagr "
                    StrSql = StrSql & " FROM empleado "
                    StrSql = StrSql & " WHERE ternro = " & rs("ternro")
                    OpenRecordset StrSql, rs2
                    If rs2.EOF Then
                        Flog.writeline "No se encontraron la fecha de alta para el empleado.ternro(" & rs("ternro") & ")"
                        FECHA_INGRESO = "00000000"
                        Antiguedad = 0
                    Else
                       If Not IsNull(rs2("empfaltagr")) Then
                        FECHA_INGRESO = Replace(CStr(Format(CDate(rs2("empfaltagr")), "ddmmyyyy")), "/", "")
                        Antiguedad = DateDiff("d", CDate(rs2("empfaltagr")), Now)
                       End If
                    End If
                Else
                    FECHA_INGRESO = Replace(CStr(Format(CDate(rs2("altfec")), "ddmmyyyy")), "/", "")
                    Antiguedad = DateDiff("d", CDate(rs2("altfec")), Now)
                End If
                rs2.Close
                
                
                'Busco CP_LABORAL
                StrSql = " SELECT codigopostal "
                StrSql = StrSql & " FROM empleado  "
                StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.tenro = 1 " 'Sucursal
                StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(l_fecestr) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(l_fecestr) & "))"
                StrSql = StrSql & " INNER JOIN sucursal on his_estructura.estrnro = sucursal.estrnro "
                StrSql = StrSql & " INNER JOIN cabdom ON sucursal.ternro = cabdom.ternro "
                StrSql = StrSql & " INNER JOIN detdom ON cabdom.domnro = detdom.domnro "
                StrSql = StrSql & " WHERE empleado.ternro = " & rs("ternro")
                OpenRecordset StrSql, rs2
                If rs2.EOF Then
                    Flog.writeline "No se encontraron el CP por domicilio de la sucursal para el empleado.ternro(" & rs("ternro") & ")"
                    rs2.Close
                    'Lo busco por codigo postal asociado a la estructura sucursal
                    StrSql = " SELECT nrocod "
                    StrSql = StrSql & " FROM empleado  "
                    StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.tenro = 1 " 'Sucursal
                    StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(l_fecestr) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(l_fecestr) & "))"
                    StrSql = StrSql & " INNER JOIN estr_cod ON his_estructura.estrnro = estr_cod.estrnro "
                    StrSql = StrSql & " INNER JOIN tipocod ON estr_cod.tcodnro = tipocod.tcodnro "
                    StrSql = StrSql & " AND tcodnom = 'UOCRA' "
                    StrSql = StrSql & " WHERE empleado.ternro = " & rs("ternro")
                    OpenRecordset StrSql, rs2
                    If rs2.EOF Then
                        CP_LABORAL = "0000"
                        Flog.writeline "No se encontraron el CP por codigo uocra de la estructura sucursal para el empleado.ternro(" & rs("ternro") & ")"
                    Else
                        CP_LABORAL = Format(rs2("nrocod"), "0000")
                    End If
                Else
                    CP_LABORAL = Format(rs2("codigopostal"), "0000") 'valores de la liquidacion :p
                End If
                CP_LABORAL = Right("0000" & CP_LABORAL, 4)
                rs2.Close
                
                'Lo busco por codigo asociado a la estructura CONVENIO
                StrSql = " SELECT nrocod "
                StrSql = StrSql & " FROM empleado  "
                StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.tenro = 19 " 'Convenio
                StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(l_fecestr) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(l_fecestr) & "))"
                StrSql = StrSql & " INNER JOIN estr_cod ON his_estructura.estrnro = estr_cod.estrnro "
                StrSql = StrSql & " INNER JOIN tipocod ON estr_cod.tcodnro = tipocod.tcodnro "
                StrSql = StrSql & " AND tcodnom = 'UOCRA' "
                StrSql = StrSql & " WHERE empleado.ternro = " & rs("ternro")
                OpenRecordset StrSql, rs2
                If rs2.EOF Then
                    CONVENIO = "00"
                    Flog.writeline "No se encontraron el CP por codigo uocra de la estructura CONVENIO para el empleado.ternro(" & rs("ternro") & ")"
                Else
                    'CONVENIO = rs2("nrocod")
                    CONVENIO = Format(CStr(rs2("nrocod")), "00")
                    CONVENIO = Right("0000" & CONVENIO, 2)
                End If
                
                'Lo busco por codigo asociado a la estructura CATEGORIA
                StrSql = " SELECT nrocod "
                StrSql = StrSql & " FROM empleado  "
                StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.tenro = 3 " 'CATEGORIA
                StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(l_fecestr) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(l_fecestr) & "))"
                StrSql = StrSql & " INNER JOIN estr_cod ON his_estructura.estrnro = estr_cod.estrnro "
                StrSql = StrSql & " INNER JOIN tipocod ON estr_cod.tcodnro = tipocod.tcodnro "
                StrSql = StrSql & " AND tcodnom = 'UOCRA' "
                StrSql = StrSql & " WHERE empleado.ternro = " & rs("ternro")
                OpenRecordset StrSql, rs2
                If rs2.EOF Then
                    CATEGORIA = "00"
                    Flog.writeline "No se encontraron el CP por codigo uocra de la estructura CATEGORIA para el empleado.ternro(" & rs("ternro") & ")"
                Else
                    CATEGORIA = rs2("nrocod")
                    CATEGORIA = Format(CATEGORIA, "00")
                    CATEGORIA = Right("0000" & CATEGORIA, 2)
                End If
                
                'Gereracion de Archivo -----------------------------------------------------
                
                Call imprimirTexto(cuit, ArchExp, 11, True)                 'CUIT
                Call imprimirTexto(AFILIADO, ArchExp, 1, True)              'AFILIADO
                Call imprimirTexto(CUOTA_SINDICAL, ArchExp, 8, False)      'CUOTA_SINDICAL
                Call imprimirTexto(FONDO_CESE_LABORAL, ArchExp, 8, False)  'FONDO_CESE_LABORAL
                Call imprimirTexto(FECHA_INGRESO, ArchExp, 8, True)         'FECHA_INGRESO
                Call imprimirTexto(CP_LABORAL, ArchExp, 4, True)            'CP_LABORAL
                Call imprimirTexto(CONVENIO, ArchExp, 2, True)              'CONVENIO
                Call imprimirTexto(CATEGORIA, ArchExp, 2, True)             'CATEGORIA
                Call imprimirTexto(Admpublica, ArchExp, 1, True)             'Administracion publica
                
                ArchExp.writeline
                
                'Actualizo el estado del proceso
                TiempoAcumulado = GetTickCount
                cantRegistros = cantRegistros - 1
                Progreso = Progreso + IncPorc
                
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
                StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
                StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
                objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
                rs.MoveNext
            Loop
            ArchExp.Close
            
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



