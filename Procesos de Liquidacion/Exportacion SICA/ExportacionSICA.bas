Attribute VB_Name = "ExpSICA"
' ---------------------------------------------------------------------------------------------
' Descripcion: Genera exportacion SICA
' Autor      : Dimatz Rafael
' Fecha      : 01/02/2016
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "01/02/2016" 'CAS-34258 - PIRAMIDE - Reporte SICA - Dimatz Rafael
'Global Const UltimaModificacion = " " 'Version Inicial

'Global Const Version = "1.01"
'Global Const FechaModificacion = "03/02/2016" 'CAS-34258 - PIRAMIDE - Reporte SICA - Dimatz Rafael
'Global Const UltimaModificacion = "Se corrigio para que cree la carpeta ExpSICA en In-Out " 'Version Inicial

'Global Const Version = "1.02"
'Global Const FechaModificacion = "05/02/2016" 'CAS-34258 - PIRAMIDE - Reporte SICA - Dimatz Rafael
'Global Const UltimaModificacion = "Se corrigio para que no salga el separador en el ultimo campo. Se agrego Fecha de Inicio " 'Version Inicial

'Global Const Version = "1.03"
'Global Const FechaModificacion = "17/03/2016" 'CAS-34258 - PIRAMIDE - Reporte SICA - Dimatz Rafael
'Global Const UltimaModificacion = "Se corrigio para que no de error el logs cuando quiere asignar el ultimo Tercero y en realidad no hay mas terceros"

'Global Const Version = "1.04"
'Global Const FechaModificacion = "22/03/2016" 'CAS-34258 - PIRAMIDE - Reporte SICA - Dimatz Rafael
'Global Const UltimaModificacion = "Se corrigio la query del Escalafon_Sica para que busque el Tipo de Codigo 203"

Global Const Version = "1.05"
Global Const FechaModificacion = "04/04/2016" 'CAS-34258 - PIRAMIDE - Reporte SICA - Dimatz Rafael
Global Const UltimaModificacion = "Se corrigio para que la Fecha de Ingreso Salga separada por - "
'-----------------------------------------------------------------------------------

Global fs, f
'Global Flog
Global NroProceso As Long

Global Path As String
Global HuboErrores As Boolean

Global EmpErrores As Boolean

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

Dim strCmdLine As String
Dim Nombre_Arch As String

Dim rs As New ADODB.Recordset
Dim PID As String
Dim Parametros As String
Dim ArrParametros As Variant
Dim arrSindicatos

  
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
    
    Nombre_Arch = PathFLog & "ExportacionSICA" & "-" & NroProceso & ".log"
    
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
    StrSql = StrSql & " AND btprcnro = " & 462
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
       'Obtengo los parametros del proceso
        Flog.writeline "Ingresa a clasificar parametros"
        Parametros = rs!bprcparam
        
        If Not IsNull(Parametros) Then
            ArrParametros = Split(Parametros, "@")
            Flog.writeline "Envía los parámetros a la función generar"
            Flog.writeline "Parametros: " & Parametros
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
        cadena = Left(CStr(Texto), Longitud)
    End If
    
    u = CInt(Longitud) - longTexto
     If u <= 0 Then
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
        longTexto = Longitud
        cadena = " "
    Else
        longTexto = Len(Texto)
        cadena = CStr(Texto)
    End If
    
    u = CInt(Longitud) - longTexto
    If u <= 0 Then
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
        
        If Len(Texto) < Longitud Then
            'cadena = String(Longitud - Len(cadena), "0") & cadena
            cadena = String(Longitud - Len(Texto), "0") & Texto
        Else
            cadena = Left(Texto, Longitud)
        End If
    
        'cadena = CStr(Texto)
    End If
    
    archivo.Write CStr(cadena)
    
End Sub
    
Sub TextoScomaYnum(Texto, archivo, Longitud, derecha)
'Rutina que imprime un nro recibido quitandole la coma y rellenando con ceros
'hasta alcanzar la longitud especificada.
'Si el nro no tiene decimales, se agregan dos ceros al final
Dim aux
Dim relleno
Dim strTexto
Dim cn
Dim VecDecim

relleno = "0"
If IsNull(Texto) Or Texto = "" Then
  cn = Longitud
  'no se pasó ningún texto
  archivo.Write String(cn, relleno)
Else
  If InStr(Texto, ".") = 0 Then
    aux = Texto & "00"
  Else
    VecDecim = Split(Texto, ".")
    If Len(VecDecim(1)) = 1 Then
      aux = Replace(Texto, ".", "")
      aux = aux & "0"
    Else
      aux = Replace(Texto, ".", "")
    End If
  End If
  cn = Longitud - Len(aux)
  If cn > 0 Then
    strTexto = String(cn, relleno) & aux
  Else
    strTexto = aux
  End If
  archivo.Write strTexto
End If
        
End Sub

Sub prntTextoCeros(Texto, archivo, Longitud, derecha)
'Imprime un texto recibido en una longitud específica, completando con ceros
'a isquierda o derecha, según se indique
Dim lonrell
Dim relle
relle = "0" 'cambiar si se necesita otro relleno hasta completar
If IsNull(Texto) Or Texto = "" Then
  lonrell = Longitud 'toda la longitud deseada con relleno
  archivo.Write String(lonrell, relle)
Else
  lonrell = Longitud - Len(Texto)
  'verificamos si rellena a derecha o izquierda
  If derecha = True Then
    archivo.Write String(lonrell, relle) & Texto
  Else
    archivo.Write Texto & String(lonrell, relle)
  End If
End If

End Sub

Private Sub Generar_Archivo(ByVal ArrParametros As Variant)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera la exportacion
' Autor      : Dimatz Rafael
' Fecha      : 26/01/2016
' Ultima Mod :
' Descripcion: Version Inicial

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

Dim Nombre_Arch As String
Dim NroModelo As Long
Dim directorio As String
Dim carpeta
Dim fs1, fs2

Dim Sep As String

Dim cantRegistros As Long
Dim Empresa As String

Dim Legajo As String
Dim Tercero As Long
Dim estrSEC As Long

Dim cuil As String '13 caracteres con guiones.
Dim Apellido As String ' 30 digitos, completa con blancos
Dim Nombre As String ' 30 digitos, completa con blancos
Dim Sexo As String
Dim f_ingreso As String ' 8 digitos formato (yyyymmdd)
Dim med_jornada As String '1 dígito (0=no, 1=si)
Dim hab_aportes As String ' 20 dígitos, completa con ceros a derecha
Dim impor_noremun As String ' 20 dígitos, completa con ceros a derecha
Dim categoria As String ' 2
Dim tip_docum As String '1 digito
Dim nro_docum As String '15 digitos, completa con blancos
Dim puesto As String '2 digitos, completa con cero a derecha
Dim licencia As String ' 2 dígitos, completa con ceros


Dim NroReporte As Long

Dim Codigo_Haber_Descuento
Dim sub_tipo_de_codigo
Dim descripcion
Dim tipo_de_codigo
Dim TipodeCod
Dim Codigo
Dim Tipo_Est
Dim Tipo_TCR
Dim Tipo_CRB
Dim Tipo_RNB
Dim Tipo_CNR
Dim Tipo_TAS
Dim Tipo_DES
Dim Situacion
Dim Monto

Dim l_Lista_ternro As String
Dim l_Lista_ternro_temp
Dim I As Long

Dim Lista_Sec As String
Dim Lista_SalRemun As String
Dim Lista_SalNremun As String

Dim l_mes
Dim l_anio
Dim Nro_ternro_ant
Dim Nro_ternro

Dim l_tipo_liq
Dim tipo_liq
Dim l_nro_liq
Dim Id_Sica
Dim Escalafon_Sica
Dim Codigo_Liq

Lista_Sec = 0
Lista_SalRemun = 0
Lista_SalNremun = 0

Dim arrSindicatos

'Escribe log de ingreso
Flog.writeline "Ingresó a generar el archivo"

Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset
    
    On Error GoTo ME_Local

        l_filtro = CStr(ArrParametros(0))
        Flog.writeline "Filtro: " & l_filtro
        'fecha estructura
        l_fecestr = C_Date(ArrParametros(1))
        'Estructura 1
        l_tenro1 = CInt(ArrParametros(2))
        l_estrnro1 = CInt(ArrParametros(3))
        'Estructura 2
        l_tenro2 = CInt(ArrParametros(4))
        l_estrnro2 = CInt(ArrParametros(5))
        'Estructura 3
        l_tenro3 = CInt(ArrParametros(6))
        l_estrnro3 = CInt(ArrParametros(7))
        'Periodo liquidacion desde
        l_pliqdesde = CLng(IIf(ArrParametros(8) = "null" Or ArrParametros(8) = "", 0, ArrParametros(8)))
        'Periodo liquidacion hasta
        l_pliqhasta = CLng(IIf(ArrParametros(9) = "null" Or ArrParametros(9) = "", 0, ArrParametros(9)))
        'fecha periodo desde
        l_desde = C_Date(ArrParametros(10))
        'fecha periodo hasta
        l_hasta = C_Date(ArrParametros(11))
        'Procesos estados
        '   Aprob.Definitivo = 3
        '   Aprob.Provisorio = 2
        '   Liquidado = 1
        '   No Liquidado = 0
        '   Todos = -1
        l_proaprob = CLng(IIf(ArrParametros(12) = "", 0, ArrParametros(12))) 'CInt(ArrParametros(12))
        ' Lista de Procesos
        l_listaproc = CStr(ArrParametros(13))
        Flog.writeline "Procesos seleccionados: " & l_listaproc
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
        
        l_tipo_liq = CStr(ArrParametros(19))
        
        If l_tipo_liq = -1 Then
            tipo_liq = "Normal"
            l_nro_liq = 0
        Else
            If l_tipo_liq = 0 Then
                tipo_liq = "Complementaria"
                l_nro_liq = 1
            Else
                tipo_liq = "Sac"
                l_nro_liq = 1
            End If
        End If
        
        l_Lista_ternro = "0"
        If l_lista_emp <> "0" Then
            l_Lista_ternro_temp = Split(l_lista_emp, ",")
            For I = 0 To UBound(l_Lista_ternro_temp) - 1
                l_Lista_ternro = l_Lista_ternro & "," & l_Lista_ternro_temp(I + 1)
                I = I + 1
            Next I
        End If
    '////////////////////////////////////////////////////////////////////
    On Error GoTo ME_Local
    
    Flog.writeline "Va a pedir el nro de modelo"
    NroModelo = 923 'Exportacion SICA
    'Directorio \ExpSICA
    'Directorio de exportacion
    StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Flog.writeline "Ingresa a buscar el directorio de salida"
        directorio = Trim(rs!sis_dirsalidas)
    End If
    
    'Selecciona el modelo de exportación
    StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
    OpenRecordset StrSql, rs_Modelo
    If Not rs_Modelo.EOF Then
      Flog.writeline "Encontró el número de modelo: " & NroModelo
        If Not IsNull(rs_Modelo!modarchdefault) Then
            directorio = directorio & Trim(rs_Modelo!modarchdefault)
            If Right(directorio, 1) <> "\" Then
                directorio = directorio & "\"
            End If
        Else
            Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta desteino. El archivo será generado en el directorio default"
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo " & NroModelo & ". El archivo será generado en el directorio default"
    End If

    'Obtengo los datos del separador
    Sep = rs_Modelo!modseparador
    'SepDec = rs_Modelo!modsepdec
    Flog.writeline Espacios(Tabulador * 0) & "Separador seleccionado: " & Sep
    
    'Seteo el nombre del archivo a exportar en el directorio indicado
    l_mes = Month(l_desde)
    If l_mes < 10 Then
        l_mes = "0" & l_mes
    End If
    l_anio = Year(l_hasta)
    
    StrSql = " SELECT nrocod FROM estr_cod "
    StrSql = StrSql & " INNER JOIN tipocod ON tipocod.tcodnro = estr_cod.tcodnro"
    StrSql = StrSql & " WHERE (tipocod.tcodnro = 202)"
    StrSql = StrSql & " AND estrnro = " & l_empresa
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Id_Sica = rs!nrocod
    Else
        Flog.writeline "No se Configuro el Tipo de Codigo 202 Id Sica"
    End If
    rs.Close
    
    StrSql = " SELECT nrocod FROM estr_cod "
    StrSql = StrSql & " INNER JOIN tipocod ON tipocod.tcodnro = estr_cod.tcodnro"
    StrSql = StrSql & " WHERE (tipocod.tcodnro = 203)"
    StrSql = StrSql & " AND estrnro = " & l_empresa
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Escalafon_Sica = rs!nrocod
    Else
        Flog.writeline "No se Configuro el Tipo de Codigo 203 Escalafon Sica"
    End If

    Nombre_Arch = directorio & Id_Sica & "_" & l_mes & "_" & l_anio & "_" & tipo_liq & "_" & l_nro_liq & ".txt"
    Flog.writeline Espacios(Tabulador * 1) & "Se crea el archivo: " & Nombre_Arch
    
    Set fs2 = CreateObject("Scripting.FileSystemObject")
    'Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    On Error Resume Next
    'If Err.Number <> 0 Then
        'Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
        Set carpeta = fs2.CreateFolder(directorio)
    'End If
    'desactivo el manejador de errores
    Set ArchExp = fs2.CreateTextFile(Nombre_Arch, True)

    On Error GoTo ME_Local

    'Configuracion del Reporte
    NroReporte = 506
    
    Lista_Sec = "0"
    Lista_SalRemun = "0"
    Lista_SalNremun = "0"
    
    'De aquí obtiene las columnas acumuladores para el reporte en confrep
    StrSql = "SELECT * FROM confrep "
    StrSql = StrSql & " WHERE repnro = " & NroReporte
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Flog.writeline "No hay Conceptos ni Acumuladores configurados "
        'Exit Sub
    Else
        Tipo_TCR = 0
        Tipo_CRB = 0
        Tipo_RNB = 0
        Tipo_CNR = 0
        Tipo_TAS = 0
        Tipo_DES = 0
        Do While Not rs.EOF
        'Por cada concepto en el reporte, se fija el tipo y acumula donde corresponda
            'Flog.writeline "Tipo en confrep: " & rs("conftipo")
            Select Case UCase(rs("conftipo"))
                Case "TE"
                    Tipo_Est = rs!confval
                Case "TCR"
                    Tipo_TCR = Tipo_TCR & "," & rs!confval
                Case "CRB"
                    Tipo_CRB = Tipo_CRB & "," & rs!confval
                Case "RNB"
                    Tipo_RNB = Tipo_RNB & "," & rs!confval
                Case "CNR"
                    Tipo_CNR = Tipo_CNR & "," & rs!confval
                Case "TAS"
                    Tipo_TAS = Tipo_TAS & "," & rs!confval
                Case "DES"
                    Tipo_DES = Tipo_DES & "," & rs!confval
            End Select
            rs.MoveNext
        Loop
    End If
    
    
    'Pone a funcionar, de los filtros seleccionados, los que se eligen en ordenamiento
    If l_tenro3 <> 0 Then ' esto ocurre solo cuando se seleccionan los tres niveles
         StrSql = " SELECT DISTINCT empleado.ternro, empleg, estact1.tenro tenro1, estact1.estrnro estrnro1, "
         StrSql = StrSql & " estact2.tenro tenro2, estact2.estrnro estrnro2, estact3.tenro tenro3, estact3.estrnro estrnro3, concepto.concnro, dlimonto "
         StrSql = StrSql & " FROM cabliq "
         If l_listaproc = "" Or l_listaproc = "0" Then
            StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
         Else
            StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & l_listaproc & ") "
         End If
         
         StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & l_tenro1
         StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(l_fecestr) & " AND (estact1.htethasta IS NULL OR estact1.htethasta>=" & ConvFecha(l_fecestr) & "))"
         If l_estrnro1 <> 0 Then 'cuando se le asigna un valor al nivel 1
            StrSql = StrSql & " AND estact1.estrnro =" & l_estrnro1
         End If
         
         StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & l_tenro2
         StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(l_fecestr) & " AND (estact2.htethasta IS NULL OR estact2.htethasta>=" & ConvFecha(l_fecestr) & "))"
         If l_estrnro2 <> 0 Then 'cuando se le asigna un valor al nivel 2
            StrSql = StrSql & " AND estact2.estrnro =" & l_estrnro2
         End If
         
         StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
         StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
         StrSql = StrSql & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro  = " & l_tenro3
         StrSql = StrSql & " AND (estact3.htetdesde<=" & ConvFecha(l_fecestr) & " AND (estact3.htethasta IS NULL OR estact3.htethasta>=" & ConvFecha(l_fecestr) & "))"
         If l_estrnro3 <> 0 Then 'cuando se le asigna un valor al nivel 3
            StrSql = StrSql & " AND estact3.estrnro =" & l_estrnro3
         End If
         
         StrSql = StrSql & " WHERE " & l_filtro
         
         If l_Lista_ternro <> "0" Then
           StrSql = StrSql & " AND empleado.ternro IN (" & l_Lista_ternro & ")"
         End If
            
    ElseIf l_tenro2 <> 0 Then  'ocurre cuando se selecciono hasta el segundo nivel de estructura
         StrSql = "SELECT DISTINCT empleado.ternro, empleg, estact1.tenro tenro1, estact1.estrnro estrnro1, "
         StrSql = StrSql & " estact2.tenro tenro2, estact2.estrnro estrnro2, concepto.concnro, dlimonto "
         StrSql = StrSql & " FROM cabliq "
         If l_listaproc = "" Or l_listaproc = "0" Then
           StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
         Else
           StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & l_listaproc & ") "
         End If
         
         StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & l_tenro1
         StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(l_fecestr) & " AND (estact1.htethasta IS NULL OR estact1.htethasta>=" & ConvFecha(l_fecestr) & "))"
         If l_estrnro1 <> 0 Then 'cuando se le asigna un valor al nivel 1
            StrSql = StrSql & " AND estact1.estrnro =" & l_estrnro1
         End If
         
         StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
         StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
         StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & l_tenro2
         StrSql = StrSql & " AND (estact2.htetdesde <= " & ConvFecha(l_fecestr) & " AND (estact2.htethasta IS NULL OR estact2.htethasta >= " & ConvFecha(l_fecestr) & "))"
         If l_estrnro2 <> 0 Then 'cuando se le asigna un valor al nivel 2
            StrSql = StrSql & " AND estact2.estrnro = " & l_estrnro2
         End If
         
         StrSql = StrSql & " WHERE " & l_filtro
         
         If l_Lista_ternro <> "0" Then
           StrSql = StrSql & " AND empleado.ternro IN (" & l_Lista_ternro & ")"
         End If
              
    ElseIf l_tenro1 <> 0 Then  ' Cuando solo seleccionamos el primer nivel de estructura
        StrSql = "SELECT DISTINCT empleado.ternro, empleg,  estact1.tenro tenro1, estact1.estrnro estrnro1, concepto.concnro, dlimonto "
        StrSql = StrSql & " FROM cabliq "
        If l_listaproc = "" Or l_listaproc = "0" Then
            StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
        Else
            StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & l_listaproc & ") "
        End If
        
        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & l_tenro1
        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
        StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
        StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(l_fecestr) & " AND (estact1.htethasta IS NULL OR estact1.htethasta>=" & ConvFecha(l_fecestr) & "))"
        If l_estrnro1 <> 0 Then 'cuando se le asigna un valor al nivel 1
            StrSql = StrSql & " AND estact1.estrnro =" & l_estrnro1
        End If
        
        StrSql = StrSql & " WHERE " & l_filtro
        
        If l_Lista_ternro <> "0" Then
           StrSql = StrSql & " AND empleado.ternro IN (" & l_Lista_ternro & ")"
        End If
        
    Else  ' cuando no hay nivel de estructura seleccionado
        StrSql = " SELECT DISTINCT empleado.ternro, empleg, concepto.concnro, dlimonto "
        StrSql = StrSql & " FROM cabliq "
        If l_listaproc = "" Or l_listaproc = "0" Then
            StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
        Else
            StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & l_listaproc & ") "
        End If
        StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro =empleado.ternro AND tenro = 10 "
        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
        StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
        StrSql = StrSql & " WHERE ( his_estructura.htetdesde <= " & ConvFecha(l_hasta) & " AND (his_estructura.htethasta >=" & ConvFecha(l_desde) & " OR his_estructura.htethasta IS NULL))"
        StrSql = StrSql & " AND " & l_filtro
        
        If l_Lista_ternro <> "0" Then
            StrSql = StrSql & " AND empleado.ternro IN (" & l_Lista_ternro & ")"
        End If
        
        StrSql = StrSql & " ORDER BY empleado.ternro ASC "
        
    End If
       Flog.writeline "Consulta general: " & StrSql
    'Busco los empleados
    OpenRecordset StrSql, rs
    
    ' _________________________________________________________________________
    Flog.writeline "  SQL para control de los empleados pertenecientes al filtro seleccionado. "
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
        
'Gereracion de Archivo Cabecera -----------------------------------------------------
                  'ID ORGANISMO
                  ArchExp.Write Id_Sica
                  ArchExp.Write Sep
                  
                  'ESCALAFON
                  ArchExp.Write Escalafon_Sica
                  ArchExp.Write Sep
                                   
                  'PERIODO DE LIQUIDACION
                  ArchExp.Write l_anio
                  ArchExp.Write "-"
                  ArchExp.Write l_mes
                  ArchExp.Write Sep
                  
                  'TIPO DE LIQUIDACION
                  ArchExp.Write tipo_liq
                  ArchExp.Write Sep
                  
                  'NUMERO DE LIQUIDACION
                  ArchExp.Write l_nro_liq
                  'ArchExp.Write Sep
                  ArchExp.writeline
        
'Fin Gereracion de Archivo Cabecera -----------------------------------------------------

'Ciclo por empleados
            Nro_ternro_ant = 0
            Nro_ternro = rs!Ternro
            Do Until rs.EOF 'Para todos los que cumplen con el filtro elegido + SEC por las dudas
                If Nro_ternro <> Nro_ternro_ant Then
                    'Busco el cuit del empleado
                    StrSql = " SELECT nrodoc "
                    StrSql = StrSql & " FROM ter_doc "
                    StrSql = StrSql & " WHERE ternro = " & rs("ternro")
                    StrSql = StrSql & " AND tidnro = 10 "
                    OpenRecordset StrSql, rs2
                    If rs2.EOF Then
                        Flog.writeline "No se encontró el cuit para el empleado.ternro(" & rs("ternro") & ")"
                        cuil = " "
                    Else
                        cuil = Replace(rs2!NroDoc, "-", "")
                    End If
                    rs2.Close
                    
                    'Busco Fecha de Ingreso
                    StrSql = " SELECT altfec "
                    StrSql = StrSql & " FROM fases "
                    StrSql = StrSql & " WHERE empleado = " & rs("ternro")
                    StrSql = StrSql & " AND estado = -1 "
                    OpenRecordset StrSql, rs2
                    If rs2.EOF Then
                        rs2.Close
                        Flog.writeline "No se encontró la fase con fecha de alta reconocida para el empleado.ternro(" & rs("ternro") & ")"
                        StrSql = " SELECT empfaltagr "
                        StrSql = StrSql & " FROM empleado "
                        StrSql = StrSql & " WHERE ternro = " & rs("ternro")
                        OpenRecordset StrSql, rs2
                        If rs2.EOF Then
                            Flog.writeline "No se encontró la fecha de alta para el empleado.ternro(" & rs("ternro") & ")"
                            f_ingreso = ""
                            rs2.Close
                        Else
                            f_ingreso = CStr(Format(CDate(rs2("empfaltagr")), "yyyy-mm-dd"))
                            rs2.Close
                        End If
                    Else
                        f_ingreso = CStr(Format(CDate(rs2("altfec")), "yyyy-mm-dd"))
                        rs2.Close
                    End If
                        
                    'traigo tipo y número de documento (tip_docum, nro_docum)
                    StrSql = "SELECT tidsigla, nrodoc"
                    StrSql = StrSql & " FROM ter_doc"
                    StrSql = StrSql & " INNER JOIN tipodocu ON tipodocu.tidnro = ter_doc.tidnro"
                    StrSql = StrSql & " WHERE ternro=" & rs("ternro")
                    OpenRecordset StrSql, rs2
                    Flog.writeline "Búsqueda de documento: " & StrSql
                    'verifico existencia de documento
                    If rs2.EOF Then
                      tip_docum = ""
                      nro_docum = 0
                    Else
                      tip_docum = rs2!tidsigla
                      nro_docum = rs2!NroDoc
                    End If
                    rs2.Close 'libero el objeto para continuar
                    
                    'traigo los datos de nombre y apellido
                    StrSql = "SELECT ternom,ternom2, terape, terape2,tersex "
                    StrSql = StrSql & "FROM tercero "
                    StrSql = StrSql & "WHERE tercero.ternro = " & rs!Ternro
                    OpenRecordset StrSql, rs2
                                          
                    If rs2.EOF Then
                      Nombre = ""
                      Apellido = ""
                    Else
                      Nombre = rs2!ternom & " " & rs2!ternom2
                      Apellido = rs2!terape & " " & rs2!terape2
                      If rs2!tersex = -1 Then
                        Sexo = "Masculino"
                      Else
                        Sexo = "Femenino"
                      End If
                    End If
                    rs2.Close
                                    
                    '------------------------------------------------------------------
                    'Busco el valor de la Categoria
                    '------------------------------------------------------------------
                    StrSql = " SELECT estrdabr "
                    StrSql = StrSql & " From his_estructura"
                    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
                    StrSql = StrSql & " AND htetdesde <= " & ConvFecha(l_hasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(l_hasta) & ") And his_estructura.tenro = 3 And his_estructura.ternro = " & rs!Ternro
                    OpenRecordset StrSql, rs2
                    
                    If Not rs2.EOF Then
                       categoria = rs2!estrdabr
                    Else
                       categoria = " "
                       Flog.writeline "Error al obtener los datos de la Categoria: " & StrSql
                    End If
                    rs2.Close
                    
                    '------------------------------------------------------------------
                    'Busco el valor de la Puesto
                    '------------------------------------------------------------------
                    StrSql = " SELECT estrdabr "
                    StrSql = StrSql & " From his_estructura"
                    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
                    StrSql = StrSql & " AND htetdesde <= " & ConvFecha(l_hasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(l_hasta) & ") And his_estructura.tenro = 4 And his_estructura.ternro = " & rs!Ternro
                    OpenRecordset StrSql, rs2
                    
                    If Not rs2.EOF Then
                       puesto = rs2!estrdabr
                    Else
                       puesto = " "
                       Flog.writeline "Error al obtener los datos del Puesto:" & StrSql
                    End If
                    rs2.Close
                
                    '------------------------------------------------------------------
                    'Busco el valor de la Situacion
                    '------------------------------------------------------------------
                    StrSql = " SELECT estrdabr "
                    StrSql = StrSql & " From his_estructura"
                    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
                    StrSql = StrSql & " AND htetdesde <= " & ConvFecha(l_hasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(l_hasta) & ") "
                    StrSql = StrSql & " AND his_estructura.tenro = " & Tipo_Est
                    StrSql = StrSql & " AND his_estructura.ternro = " & rs!Ternro
                    OpenRecordset StrSql, rs2
                    
                    If Not rs2.EOF Then
                       Situacion = rs2!estrdabr
                    Else
                       Situacion = " "
                       Flog.writeline "Error al obtener los datos de la Situacion: " & StrSql
                    End If
                    rs2.Close
                End If
                                   
                sub_tipo_de_codigo = ""
                descripcion = ""
                tipo_de_codigo = ""
                
                StrSql = " SELECT conctexto,concabr,conccod,concepto.tconnro "
                StrSql = StrSql & "FROM concepto "
                StrSql = StrSql & "INNER JOIN tipconcep ON tipconcep.tconnro=concepto.tconnro "
                StrSql = StrSql & "WHERE concnro = " & rs!ConcNro
                OpenRecordset StrSql, rs2

                If Not rs2.EOF Then
                    Codigo_Haber_Descuento = rs2!ConcCod
                    descripcion = rs2!concabr
                    tipo_de_codigo = rs2!tconnro
                    sub_tipo_de_codigo = rs2!Conctexto
               End If
               rs2.Close
               tipo_de_codigo = Trim(Str(tipo_de_codigo))
               TipodeCod = InStr(Tipo_TCR, tipo_de_codigo)
               'Flog.writeline "Tipo de Codigo: " & TipodeCod
               If TipodeCod <> 0 Then
                    Codigo = "Remunerativo"
               Else
                    TipodeCod = InStr(Tipo_CRB, tipo_de_codigo)
                    If TipodeCod <> 0 Then
                         Codigo = "Remunerativo Bonificable"
                         Flog.writeline "El Codigo es: " & tipo_de_codigo & " El Tipo de Codigo es: " & Codigo
                    Else
                         TipodeCod = InStr(Tipo_RNB, tipo_de_codigo)
                         If TipodeCod <> 0 Then
                            Codigo = "Remunerativo No Bonificable"
                            Flog.writeline "El Codigo es: " & tipo_de_codigo & " El Tipo de Codigo es: " & Codigo
                         Else
                            TipodeCod = InStr(Tipo_CNR, tipo_de_codigo)
                            If TipodeCod <> 0 Then
                               Codigo = "No Remunerativo No Bonificable"
                               Flog.writeline "El Codigo es: " & tipo_de_codigo & " El Tipo de Codigo es: " & Codigo
                            Else
                                TipodeCod = InStr(Tipo_TAS, tipo_de_codigo)
                                If TipodeCod <> 0 Then
                                   Codigo = "Adicionales Sociales"
                                   Flog.writeline "El Codigo es: " & tipo_de_codigo & " El Tipo de Codigo es: " & Codigo
                                Else
                                    TipodeCod = InStr(Tipo_DES, tipo_de_codigo)
                                    If TipodeCod <> 0 Then
                                       Codigo = "Descuentos"
                                        Flog.writeline "El Codigo es: " & tipo_de_codigo & " El Tipo de Codigo es: " & Codigo
                                    Else
                                        Flog.writeline "No se encontro un Tipo de Codigo configurado"
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                    
               'Monto
                Monto = rs!dlimonto
'Gereracion de Archivo Cuerpo -----------------------------------------------------
                  'Cuil
                  ArchExp.Write cuil
                  ArchExp.Write Sep
                  'Numero de Documento
                  ArchExp.Write nro_docum
                  ArchExp.Write Sep
                  'Tipo de Documento
                  ArchExp.Write tip_docum
                  ArchExp.Write Sep
                  'Nombre
                  ArchExp.Write Nombre
                  ArchExp.Write Sep
                  'Apellido
                  ArchExp.Write Apellido
                  ArchExp.Write Sep
                  'Sexo
                  ArchExp.Write Sexo
                  ArchExp.Write Sep
                  'ID ORGANISMO
                  ArchExp.Write Id_Sica
                  ArchExp.Write Sep
                  'Categoria
                  ArchExp.Write categoria
                  ArchExp.Write Sep
                  'Puesto
                  ArchExp.Write puesto
                  ArchExp.Write Sep
                  'Situacion
                  ArchExp.Write Situacion
                  ArchExp.Write Sep
                  'Fecha de Ingreso
                  ArchExp.Write f_ingreso
                  ArchExp.Write Sep
                  'Codigo de Liquidacion 0 por Defecto
                  Codigo_Liq = 0
                  ArchExp.Write Codigo_Liq
                  ArchExp.Write Sep
                  'Codigo de Haber Descuento
                  ArchExp.Write Codigo_Haber_Descuento
                  ArchExp.Write Sep
                  'Descripcion
                  ArchExp.Write descripcion
                  ArchExp.Write Sep
                  'Tipo de Codigo
                  ArchExp.Write Codigo
                  ArchExp.Write Sep
                  'Sub Tipo de Codigo
                  ArchExp.Write Trim(sub_tipo_de_codigo)
                  ArchExp.Write Sep
                  'Monto
                  ArchExp.Write Monto
                  'ArchExp.Write Sep

                  ArchExp.writeline
'Fin Gereracion de Archivo Cuerpo -----------------------------------------------------
                
                'Actualizo el estado del proceso
                TiempoAcumulado = GetTickCount
                cantRegistros = cantRegistros - 1
                Progreso = Progreso + IncPorc
                
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Replace(Progreso, ",", ".")
                StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
                StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
                objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
                Nro_ternro_ant = Nro_ternro
                rs.MoveNext
                If Not rs.EOF Then
                    Nro_ternro = rs!Ternro
                End If
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



