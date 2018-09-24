Attribute VB_Name = "ExpSecRos"
' ---------------------------------------------------------------------------------------------
' Descripcion: Genera exportacion SEC Rosario
' Autor      : Verónica Bogado
' Fecha      : 09/09/2010
' Ultima Mod : 29/11/2010
' Descripcion: Cambios de especificación. Se cambió la categoría del empleado a un
' valor fijo de 20, al igual que el puesto, que pasa a ser 4 Fijo.
' ---------------------------------------------------------------------------------------------
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "09/09/2010"
'Global Const UltimaModificacion = " " 'Version Inicial

'Global Const Version = "1.01"
'Global Const FechaModificacion = "06/10/2010"
'Global Const UltimaModificacion = "Eliminar verificación de asociación a estructura SEC. El proceso se define para todos los empleados, sean o no afiliados a comercio."

'Global Const Version = "1.02"
'Global Const FechaModificacion = "08/02/2011"
'Global Const UltimaModificacion = "Arreglo de fecha de baja faltante"

'Global Const Version = "1.03"
'Global Const FechaModificacion = "05/10/2011" ' Zamarbide Juan Alberto - CAS-13998 - HORWATH LITORAL - BUG Reporte SEC para importacion de Empleados
'Global Const UltimaModificacion = "" '------------------------------ CUSTOM HORWATH LITORAL -------------------------------------
                                     'Se comentaron un par de líneas para que verifique si el empleado tiene o no convenio.
                                     'Se corrigió la forma de traer la Categoría, ya que tenía la estructura de Teletech
                                     'Se agregó la lógica para traer el Puesto y las Licencias.
                                     
'Global Const Version = "1.04"
'Global Const FechaModificacion = "04/11/2011" ' Zamarbide Juan Alberto - CAS-13998 - HORWATH LITORAL - BUG Reporte SEC para importacion de Empleados
'Global Const UltimaModificacion = "" 'Se corrigió la media jornada/completa, por que siempre devolvía 0

'Global Const Version = "1.05"
'Global Const FechaModificacion = "17/10/2012" ' CAS-17249 - HEIDT- INSTALACION DEL SEC EN TEST80
'Global Const UltimaModificacion = "" 'Se agrego el separador de directorio al archivo de log

'Global Const Version = "1.06"
'Global Const FechaModificacion = "24/10/2012" ' CAS-17280- HORWATH LITORAL- ERROR EN SEC - Deluchi Ezequiel
'Global Const UltimaModificacion = "" 'Se corrigio las consultas que traen puesto y categoria.
                                     'En lugar de trear la fecha de alta reconocida, ahora trae la fecha de alta de la ultima fecha activa.

'Global Const Version = "1.07"
'Global Const FechaModificacion = "06/11/2012" ' CAS-17280- HORWATH LITORAL- ERROR EN SEC - Deluchi Ezequiel
'Global Const UltimaModificacion = "" 'Se corrigio las consultas que traen sindicato, para ver si es afiliado o no.
                                     
'Global Const Version = "1.08"
'Global Const FechaModificacion = "23/01/2013" ' CAS 18146 - Crowe Horwath - Error Reporte Exportacion SEC Rosario - Carmen Quintero
'Global Const UltimaModificacion = "" 'Se modificó la consulta que busca el puesto del empleado.
                                     'Se corrigió caso cuando el empleado tiene más de un puesto o situación de revista en el mes.
'Global Const Version = "1.09"
'Global Const FechaModificacion = "26/03/2013" ' CAS-18146 - Crowe Horwath - Error Reporte Exportacion SEC Rosario-(CAS-15298) - Carmen Quintero
'Global Const UltimaModificacion = "" 'Se modificó la consulta que busca los empleados seleccionados por el filtro.
                                     
Global Const Version = "1.10"
Global Const FechaModificacion = "27/01/2016" ' CAS-35004 - Monasterio base 3 - Bug en puesto reporte SEC-Miriam Ruiz
Global Const UltimaModificacion = "" 'Se corrige puesto y situación de revista.
                                     
                                                      

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
Global fs, f
'Global Flog
Global NroProceso As Long

Global Path As String
Global HuboErrores As Boolean

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
    
    Nombre_Arch = PathFLog & "ExportacionSECRosario" & "-" & NroProceso & ".log"
    
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
    StrSql = StrSql & " AND btprcnro = " & 272
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
' Autor      : Verónica Bogado
' Fecha      : 09/09/2010
' Ultima Mod : 04/03/2011
' Descripcion: Errores cuando levanta parámetros de fecha oracle

'CUIL                       13
'APELLIDO                   30
'NOMBRE                     30
'FECHA INGRESO              8 (YYYYMMAA)
'MEDIA JORNADA              1 (0= NO, 1=SI)
'HABERES C_APORTE           20 SIN SEP DECIMAL, COMPLETA CERO A DERECHA
'IMPORTE NO REMUN           20 SIN SEP DECIMAL, COMPLETA CERO A DERECHA
'CATEGORIA                  2 NUMERICO
'ES AFILIADO                1 BIT (0=NO, 1=SI)
'FECHA BAJA                 8 (YYYYMMAA)(COMPLETA CON 0 SI NO HAY)
'TIPO DOCUMENTO             1
'NRO DOCUMENTO              15 COMPLETA CON BLANCOS
'PUESTO                     2 COMPLETA CON CERO A DERECHA
'LICENCIA                   2 COMPLETA CON CERO A DERECHA Y SI NO HAY

'SEPARADOR DE TODOS LOS CAMPOS: ";"
'SEPARADOR DECIMAL: No usa

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
Dim fs1, fs2

Dim Sep As String
Dim SepDec As String

Dim cantRegistros As Long
Dim Empresa As String

Dim Legajo As String
Dim Tercero As Long
Dim estrSEC As Long

'CUIL                       13
'APELLIDO                   30
'NOMBRE                     30
'FECHA INGRESO              8 (YYYYMMDD)
'MEDIA JORNADA              1 (0= NO, 1=SI)
'HABERES C_APORTE           20 SIN SEP DECIMAL, COMPLETA CERO A DERECHA
'IMPORTE NO REMUN           20 SIN SEP DECIMAL, COMPLETA CERO A DERECHA
'CATEGORIA                  2 NUMERICO
'ES AFILIADO                1 BIT (0=NO, 1=SI)
'FECHA BAJA                 8 (YYYYMMDD)(COMPLETA CON 0 SI NO HAY)
'TIPO DOCUMENTO             1
'NRO DOCUMENTO              15 COMPLETA CON BLANCOS
'PUESTO                     2 COMPLETA CON CERO A DERECHA
'LICENCIA                   2 COMPLETA CON CERO A DERECHA Y SI NO HAY

Dim cuil As String '13 caracteres con guiones.
Dim Apellido As String ' 30 digitos, completa con blancos
Dim Nombre As String ' 30 digitos, completa con blancos
Dim f_ingreso As String ' 8 digitos formato (yyyymmdd)
Dim med_jornada As String '1 dígito (0=no, 1=si)
Dim hab_aportes As String ' 20 dígitos, completa con ceros a derecha
Dim impor_noremun As String ' 20 dígitos, completa con ceros a derecha
Dim categoria As String ' 2
Dim es_afiliado As String '1 bit (0=no, 1=si)
Dim f_baja As String '8 digitos formato (yyyymmdd)
Dim tip_docum As String '1 digito
Dim nro_docum As String '15 digitos, completa con blancos
Dim puesto As String '2 digitos, completa con cero a derecha
Dim licencia As String ' 2 dígitos, completa con ceros


Dim NroReporte As Long
Dim Acumulador As Long

Dim l_acremun As Double
Dim l_acnoremun As Double
Dim l_cuil As String
Dim l_tipdoc As String
Dim l_nrodoc As String
Dim l_apellido As String
Dim l_nombre As String
Dim l_fechaingreso As String
Dim l_mediajor As Boolean
Dim l_categ As String
Dim l_afil As Boolean
Dim l_fechabaja As String

Dim l_Lista_ternro As String
Dim l_Lista_ternro_temp
Dim I As Long

Dim Lista_Sec As String
Dim Lista_SalRemun As String
Dim Lista_SalNremun As String
Dim ListaMediajor As String
Dim ListaAfiliado As String

Lista_Sec = 0
Lista_SalRemun = 0
Lista_SalNremun = 0
ListaMediajor = 0
ListaAfiliado = 0

Dim arrSindicatos

'Dim formatofecha As String

'Escribe log de ingreso
Flog.writeline "Ingresó a generar el archivo"
'formatofecha = "YYYMMDD"

Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs_Aux As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
    
    On Error GoTo ME_Local
    
    '////////////////////////////////////////////////////////////////////
       ' filtro empleg - (0 = 0) And (empleg >= 1) And (empleg <= 9999999999#)
        l_filtro = CStr(ArrParametros(0))
        Flog.writeline "Filtro: " & l_filtro
        'StrSql2 = l_filtro
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
    NroModelo = 327 'lichok xxx
    'Directorio de exportacion
    StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Flog.writeline "Ingresa a buscar el directorio de salida"
        directorio = Trim(rs!sis_dirsalidas)
    End If
    'Directorio = "C:\Export\SEC"
    
    'Selecciona el modelo de exportación
    StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
    OpenRecordset StrSql, rs_Modelo
    If Not rs_Modelo.EOF Then
      Flog.writeline "Encontró el número de modelo"
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
    SepDec = rs_Modelo!modsepdec
    Flog.writeline Espacios(Tabulador * 0) & "Separador seleccionado: " & Sep
    
    'Seteo el nombre del archivo a exportar en el directorio indicado
    Nombre_Arch = directorio & "ExpSecRos" & "-" & NroProceso & ".txt"
    Flog.writeline Espacios(Tabulador * 1) & "Se crea el archivo: " & Nombre_Arch
    
    'Set fs = CreateObject("Scripting.FileSystemObject")
    Set fs2 = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
        Set carpeta = fs1.CreateFolder(directorio)
    End If
    'desactivo el manejador de errores
    Set ArchExp = fs.CreateTextFile(Nombre_Arch, True)

    On Error GoTo ME_Local

    'Configuracion del Reporte
    NroReporte = 288
    
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
        Do While Not rs.EOF
        'Por cada concepto en el reporte, se fija el tipo y acumula donde corresponda
            Flog.writeline "Tipo en confrep: " & rs("conftipo")
            Select Case UCase(rs("conftipo"))
                Case "SEC"
                    Lista_Sec = Lista_Sec & "," & rs!confval
                Case "AC"
                    If UCase(rs("confetiq")) = "REM" Then 'Remunerativo
                        Lista_SalRemun = Lista_SalRemun & "," & rs!confval
                    ElseIf UCase(rs("confetiq")) = "NRE" Then 'No Remunerativo
                        Lista_SalNremun = Lista_SalNremun & "," & rs!confval
                    Else
                        Flog.writeline " Mal configurado el Acumulador " & rs("confetiq")
                    End If
                Case "EST"
                  If UCase(rs("confetiq")) = "MJ" Then
                    ListaMediajor = ListaMediajor & "," & rs!confval
                  ElseIf UCase(rs("confetiq")) = "AFIL" Then
                    ListaAfiliado = ListaAfiliado & "," & rs!confval
                  End If
            End Select
            rs.MoveNext
        Loop
    End If
    
    
    'Pone a funcionar, de los filtros seleccionados, los que se eligen en ordenamiento
    If l_tenro3 <> 0 Then ' esto ocurre solo cuando se seleccionan los tres niveles
'        Comentado por Carmen Quintero 26/03/2013
'        StrSql = " SELECT DISTINCT empleado.ternro "
'        StrSql = StrSql & ", estact1.tenro tenro1, estact1.estrnro estrnro1 "
'        StrSql = StrSql & ", estact2.tenro tenro2, estact2.estrnro estrnro2 "
'        StrSql = StrSql & ", estact3.tenro tenro3, estact3.estrnro estrnro3 "
'        StrSql = StrSql & " FROM empleado "
'        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro AND estact1.tenro  = " & l_tenro1
'        StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(l_fecestr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(l_fecestr) & "))"
'        If l_estrnro1 <> 0 Then 'cuando se le asigna un valor al nivel 1
'                StrSql = StrSql & " AND estact1.estrnro = " & l_estrnro1
'        End If
'        StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & l_tenro2
'        StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(l_fecestr) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(l_fecestr) & "))"
'        If l_estrnro2 <> 0 Then 'cuando se le asigna un valor al nivel 2
'            StrSql = StrSql & " AND estact2.estrnro =" & l_estrnro2
'        End If
'        StrSql = StrSql & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro  = " & l_tenro3
'        StrSql = StrSql & " AND (estact3.htetdesde<=" & ConvFecha(l_fecestr) & " AND (estact3.htethasta is null or estact3.htethasta>=" & ConvFecha(l_fecestr) & "))"
'        If l_estrnro3 <> 0 Then ' cuando se le asigna un valor al nivel 3
'                StrSql = StrSql & " AND estact3.estrnro =" & l_estrnro3
'        End If
'        StrSql = StrSql & " WHERE " & l_filtro
'        'StrSql = StrSql & " AND his_estructura.estrnro IN (" & Lista_Sec & ") "
'        If l_Lista_ternro <> "0" Then
'            StrSql = StrSql & " AND empleado.ternro IN (" & l_Lista_ternro & ")"
'        End If

         StrSql = " SELECT DISTINCT empleado.ternro, empleg, estact1.tenro tenro1, estact1.estrnro estrnro1, "
         StrSql = StrSql & " estact2.tenro tenro2, estact2.estrnro estrnro2, estact3.tenro tenro3, estact3.estrnro estrnro3 "
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
'        Comentado por Carmen Quintero 26/03/2013
'        StrSql = "SELECT DISTINCT empleado.ternro, empleg "
'        StrSql = StrSql & ", estact1.tenro tenro1, estact1.estrnro estrnro1 "
'        StrSql = StrSql & ", estact2.tenro tenro2, estact2.estrnro estrnro2 "
'        StrSql = StrSql & " FROM empleado  "
'        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro AND estact1.tenro = " & l_tenro1
'        StrSql = StrSql & " AND (estact1.htetdesde <= " & ConvFecha(l_fecestr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(l_fecestr) & "))"
'        If l_estrnro1 <> 0 Then
'            StrSql = StrSql & " AND estact1.estrnro =" & l_estrnro1
'        End If
'        StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & l_tenro2
'        StrSql = StrSql & " AND (estact2.htetdesde <= " & ConvFecha(l_fecestr) & " AND (estact2.htethasta is null or estact2.htethasta >= " & ConvFecha(l_fecestr) & "))"
'        If l_estrnro2 <> 0 Then
'            StrSql = StrSql & " AND estact2.estrnro = " & l_estrnro2
'        End If
'        StrSql = StrSql & " WHERE " & l_filtro
'        'StrSql = StrSql & " AND his_estructura.estrnro IN (" & Lista_Sec & ") "
'        If l_Lista_ternro <> "0" Then
'            StrSql = StrSql & " AND empleado.ternro IN (" & l_Lista_ternro & ")"
'        End If

         StrSql = "SELECT DISTINCT empleado.ternro, empleg, estact1.tenro tenro1, estact1.estrnro estrnro1, "
         StrSql = StrSql & " estact2.tenro tenro2, estact2.estrnro estrnro2 "
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
         StrSql = StrSql & " AND (estact2.htetdesde <= " & ConvFecha(l_fecestr) & " AND (estact2.htethasta IS NULL OR estact2.htethasta >= " & ConvFecha(l_fecestr) & "))"
         If l_estrnro2 <> 0 Then 'cuando se le asigna un valor al nivel 2
            StrSql = StrSql & " AND estact2.estrnro = " & l_estrnro2
         End If
         
         StrSql = StrSql & " WHERE " & l_filtro
         
         If l_Lista_ternro <> "0" Then
           StrSql = StrSql & " AND empleado.ternro IN (" & l_Lista_ternro & ")"
         End If
              
    ElseIf l_tenro1 <> 0 Then  ' Cuando solo seleccionamos el primer nivel de estructura
'        Comentado por Carmen Quintero 26/03/2013
'        StrSql = "SELECT DISTINCT empleado.ternro, empleg "
'        StrSql = StrSql & ", estact1.tenro tenro1, estact1.estrnro estrnro1 "
'        StrSql = StrSql & " FROM empleado  "
'        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro AND estact1.tenro = " & l_tenro1
'        StrSql = StrSql & " AND (estact1.htetdesde <= " & ConvFecha(l_fecestr) & " AND (estact1.htethasta is null or estact1.htethasta >= " & ConvFecha(l_fecestr) & "))"
'        'StrSql = StrSql & " AND his_estructura.estrnro IN (" & Lista_Sec & ") "
'        If l_estrnro1 <> 0 Then
'            StrSql = StrSql & " AND estact1.estrnro = " & l_estrnro1
'        End If
'        StrSql = StrSql & " WHERE " & l_filtro
'        'StrSql = StrSql & " AND his_estructura.estrnro IN (" & Lista_Sec & ") "
'        If l_Lista_ternro <> "0" Then
'            StrSql = StrSql & " AND empleado.ternro IN (" & l_Lista_ternro & ")"
'        End If

        StrSql = "SELECT DISTINCT empleado.ternro, empleg,  estact1.tenro tenro1, estact1.estrnro estrnro1 "
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
        
        StrSql = StrSql & " WHERE " & l_filtro
        
        If l_Lista_ternro <> "0" Then
           StrSql = StrSql & " AND empleado.ternro IN (" & l_Lista_ternro & ")"
        End If
        
    Else  ' cuando no hay nivel de estructura seleccionado
'        Comentado por Carmen Quintero 26/03/2013
'        StrSql = " SELECT DISTINCT empleado.ternro, empleg "
'        StrSql = StrSql & " FROM empleado "
'        StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro=his_estructura.ternro"
'        StrSql = StrSql & " WHERE  " & l_filtro
'        'StrSql = StrSql & " AND his_estructura.estrnro IN (" & Lista_Sec & ") "
'        If l_Lista_ternro <> "0" Then
'            StrSql = StrSql & " AND empleado.ternro IN (" & l_Lista_ternro & ")"
'        End If
        
        StrSql = " SELECT DISTINCT empleado.ternro, empleg  "
        StrSql = StrSql & " FROM cabliq "
        If l_listaproc = "" Or l_listaproc = "0" Then
            StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
        Else
            StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & l_listaproc & ") "
        End If
        
        StrSql = StrSql & " WHERE " & l_filtro
        
        If l_Lista_ternro <> "0" Then
            StrSql = StrSql & " AND empleado.ternro IN (" & l_Lista_ternro & ")"
        End If
                
        
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
            
            Do Until rs.EOF 'Para todos los que cumplen con el filtro elegido + SEC por las dudas
                
                
                '-Acumula el salario remunerativo----------------
                StrSql = " SELECT sum(almonto) Remu "
                StrSql = StrSql & " FROM empleado "
                StrSql = StrSql & " INNER JOIN cabliq ON empleado.ternro = cabliq.empleado "
                StrSql = StrSql & " INNER JOIN acu_liq  ON acu_liq.cliqnro  = cabliq.cliqnro"
                StrSql = StrSql & " WHERE acu_liq.acunro IN (" & Lista_SalRemun & ")"
                StrSql = StrSql & " AND cabliq.pronro IN (" & l_listaproc & " ) "
                StrSql = StrSql & " AND empleado.ternro = " & rs("ternro")
                OpenRecordset StrSql, rs2
                If rs2.EOF Then
                    Flog.writeline "No se encontró acumulador remunerativo para el empleado.ternro(" & rs("ternro") & ")"
                Else
                    If rs2("Remu") <> "" Then
                      hab_aportes = rs2("Remu")
                    Else
                      hab_aportes = 0
                    End If
                End If
                rs2.Close
                
                '-Acumula el salario no remunerativo -------------------
                StrSql = " SELECT sum(almonto) NoRemu "
                StrSql = StrSql & " FROM empleado "
                StrSql = StrSql & " INNER JOIN cabliq ON empleado.ternro = cabliq.empleado "
                StrSql = StrSql & " INNER JOIN acu_liq  ON acu_liq.cliqnro  = cabliq.cliqnro"
                StrSql = StrSql & " WHERE acu_liq.acunro IN (" & Lista_SalNremun & ")"
                StrSql = StrSql & " AND cabliq.pronro IN (" & l_listaproc & ")"
                StrSql = StrSql & " AND empleado.ternro = " & rs("ternro")
                OpenRecordset StrSql, rs2
                If rs2.EOF Then
                    Flog.writeline "No se encontró acumulador remunerativo para el empleado.ternro(" & rs("ternro") & ")"
                Else
                  If rs2("NoRemu") <> "" Then
                    impor_noremun = rs2("NoRemu")
                  Else
                    impor_noremun = 0
                  End If
                End If
                rs2.Close
                
                
                '**************************************************************
          
                'Busco el cuit del empleado
                StrSql = " SELECT nrodoc "
                StrSql = StrSql & " FROM ter_doc "
                StrSql = StrSql & " WHERE ternro = " & rs("ternro")
                StrSql = StrSql & " AND tidnro = 10 "
                OpenRecordset StrSql, rs2
                If rs2.EOF Then
                    Flog.writeline "No se encontró el cuit para el empleado.ternro(" & rs("ternro") & ")"
                    cuil = "00-00000000-0"
                Else
                    cuil = Left(CStr(rs2("nrodoc")), 13)
                    'CUIT = Replace(CStr(CUIT), "-", "")
                    'CUIT = Right("00000000" & CUIT, 11)
                    'CUIT = Left(CStr(CUIT), 11)
                End If
                rs2.Close
                
                'Lo busco por condición de afiliación a SEC
                es_afiliado = "0"
                StrSql = " SELECT estrnro "
                StrSql = StrSql & " FROM empleado  "
                StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.tenro = 16 " 'SINDICATO
                StrSql = StrSql & " AND ((his_estructura.htetdesde <= " & ConvFecha(l_desde) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(l_hasta) & " or his_estructura.htethasta >= " & ConvFecha(l_desde) & " ))"
                StrSql = StrSql & " OR (his_estructura.htetdesde >= " & ConvFecha(l_desde) & " AND (his_estructura.htetdesde <= " & ConvFecha(l_hasta) & ")) ) "
                StrSql = StrSql & " WHERE empleado.ternro = " & rs("ternro")
                StrSql = StrSql & " AND his_estructura.estrnro In (" & ListaAfiliado & ")"
                'Flog.Writeline "Estructura: " & StrSql
                OpenRecordset StrSql, rs2
                If rs2.EOF Then
                    es_afiliado = "0"
                    Flog.writeline "No se encontró SEC en estructura para el empleado.ternro(" & rs("ternro") & ")"
                Else
                    'arrSindicatos = Split(Lista_Sec, ",") Comentado ver 1.03
                    'For I = 0 To UBound(arrSindicatos) Comentado ver 1.03
                        'If CLng(rs2("estrnro")) = CLng(arrSindicatos(I)) Then Comentado ver 1.03
                            es_afiliado = "1"
                        'End If Comentado ver 1.03
                    'Next Comentado ver 1.03
                End If
                
               rs2.Close
               
                'Busco la fecha de alta y baja
                'Antiguedad = 0
                f_ingreso = "00000000"
                StrSql = " SELECT altfec, bajfec "
                StrSql = StrSql & " FROM fases "
                StrSql = StrSql & " WHERE empleado = " & rs("ternro")
                StrSql = StrSql & " AND estado = -1 "
                OpenRecordset StrSql, rs2
                If rs2.EOF Then
                    rs2.Close
                    Flog.writeline "No se encontró la fase con fecha de alta reconocida para el empleado.ternro(" & rs("ternro") & ")"
                    StrSql = " SELECT empfaltagr, empfecbaja "
                    StrSql = StrSql & " FROM empleado "
                    StrSql = StrSql & " WHERE ternro = " & rs("ternro")
                    OpenRecordset StrSql, rs2
                    If rs2.EOF Then
                        Flog.writeline "No se encontró la fecha de alta para el empleado.ternro(" & rs("ternro") & ")"
                        f_ingreso = "00000000"
                        f_baja = "00000000"
                        'Antiguedad = 0
                        rs2.Close
                    Else
                        f_ingreso = Replace(CStr(Format(CDate(rs2("empfaltagr")), "yyyymmdd")), "/", "")
                        'La fecha de baja puede, y es común, que no exista.
                        If Not EsNulo(rs2("empfecbaja")) Then
                          f_baja = Replace(CStr(Format(CDate(rs2("empfecbaja")), "yyyymmdd")), "/", "")
                        Else
                          f_baja = "00000000"
                        End If
                        'Antiguedad = DateDiff("d", CDate(rs2("empfaltagr")), Now)
                        rs2.Close
                    End If
                Else
                    f_ingreso = Replace(CStr(Format(CDate(rs2("altfec")), "yyyymmdd")), "/", "")
                    If Not EsNulo(rs2("bajfec")) Then
                      f_baja = Replace(CStr(Format(CDate(rs2("bajfec")), "yyyymmdd")), "/", "")
                    Else
                      f_baja = "00000000"
                    End If
                    'Antiguedad = DateDiff("d", CDate(rs2("altfec")), Now)
                    rs2.Close
                End If
                'rs2.Close
                
                'Encuentro la CATEGORIA a la que pertenece el empleado
                'Elimina la búsqueda de categoría por pedido del cliente. Descomentado y modificado ver 1.03
                StrSql = " SELECT his_estructura.estrnro, estr_cod.nrocod FROM empleado "
                StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.tenro = 3 " 'CATEGORIA
                StrSql = StrSql & " AND ((his_estructura.htetdesde <= " & ConvFecha(l_desde) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(l_hasta) & " or his_estructura.htethasta >= " & ConvFecha(l_desde) & " ))"
                StrSql = StrSql & " OR (his_estructura.htetdesde >= " & ConvFecha(l_desde) & " AND (his_estructura.htetdesde <= " & ConvFecha(l_hasta) & ")) ) "
                StrSql = StrSql & " INNER JOIN estr_cod ON his_estructura.estrnro = estr_cod.estrnro "
                StrSql = StrSql & " INNER JOIN tipocod ON estr_cod.tcodnro = tipocod.tcodnro "
                StrSql = StrSql & " WHERE empleado.ternro = " & rs("ternro")
                StrSql = StrSql & " AND tipocod.tcodcodext = '155'" 'Código para definir la categoría - Agregado ver 1.03
                OpenRecordset StrSql, rs2
                If rs2.EOF Then
                   categoria = "00"
                   Flog.writeline "No se encontró categoría para el empleado .ternro(" & rs("ternro") & ")"
                Else
                    categoria = rs2("nrocod") '------------ Agregado ver 1.03
                    'Select Case rs2("estrnro")  ---------------- Comentado ver 1.03 desde acá
                    '  Case 804: categoria = 10
                    '  Case 808: categoria = 20
                    '  Case 810: categoria = 30
                    '  Case 1235: categoria = 40
                    '  Case 1275: categoria = 20
                    '  Case 1311: categoria = 10
                    '  Case Else: categoria = "00" --------------- Comentado ver 1.03 hasta acá
                    'End Select
                End If
                rs2.Close
                
                '******************La categoría queda fija por pedido en 20 -- Comentado ver 1.03
                'categoria = 20 Comentado ver 1.03
                
                'traigo tipo y número de documento (tip_docum, nro_docum)
                StrSql = "select tidnro, nrodoc from ter_doc where ternro=" & rs("ternro")
                OpenRecordset StrSql, rs2
                Flog.writeline "Búsqueda de documento: " & StrSql
                'verifico existencia de documento
                If rs2.EOF Then
                  tip_docum = "0"
                  nro_docum = "Undefined"
                Else
                  tip_docum = rs2!tidnro
                  nro_docum = rs2!NroDoc
                End If
                rs2.Close 'libero el objeto para continuar
                
                'traigo los datos de nombre y apellido
                StrSql = "SELECT ternom,ternom2, terape, terape2 "
                StrSql = StrSql & "FROM empleado "
                StrSql = StrSql & "WHERE empleado.ternro = " & rs("ternro")
                OpenRecordset StrSql, rs2
                         
                 
                'verifico que tenga datos
                If rs2.EOF Then
                  Nombre = ""
                  Apellido = ""
                Else
                  Nombre = rs2("ternom") & " " & rs2("ternom2")
                  Apellido = rs2("terape") & " " & rs2("terape2")
                End If
                rs2.Close
                
                '----Rutina para averiguar media jornada
                med_jornada = 0
                StrSql = " SELECT his_estructura.estrnro FROM empleado "
                StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro "
                StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(l_fecestr) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(l_fecestr) & "))"
                StrSql = StrSql & " WHERE empleado.ternro = " & rs("ternro") & " and his_estructura.estrnro in (" & ListaMediajor & ")"
                OpenRecordset StrSql, rs2
                
                If Not rs2.EOF Then
                  med_jornada = 1
                End If
                rs2.Close
                
                '---------------------Puesto - - Agregado ver 1.03 --------------------
'                StrSql = " SELECT his_estructura.estrnro, estr_cod.nrocod FROM empleado "
'                StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.tenro = 4 " 'PUESTO
'                StrSql = StrSql & " AND ((his_estructura.htetdesde <= " & ConvFecha(l_desde) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(l_hasta) & " or his_estructura.htethasta >= " & ConvFecha(l_desde) & " ))"
'                StrSql = StrSql & " OR (his_estructura.htetdesde >= " & ConvFecha(l_desde) & " AND (his_estructura.htetdesde <= " & ConvFecha(l_hasta) & ")) ) "
'                StrSql = StrSql & " INNER JOIN estr_cod ON his_estructura.estrnro = estr_cod.estrnro "
'                StrSql = StrSql & " INNER JOIN tipocod ON estr_cod.tcodnro = tipocod.tcodnro "
'                StrSql = StrSql & " WHERE empleado.ternro = " & rs("ternro")
'                StrSql = StrSql & " AND tipocod.tcodcodext = '155'" 'Código para definir el puesto - Agregado ver 1.03
                              
'                StrSql = "SELECT DISTINCT his_estructura.estrnro, estr_cod.nrocod FROM empleado "
'                StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.tenro = 4 " 'Puesto
'                StrSql = StrSql & " AND (his_estructura.htetdesde >= " & ConvFecha(l_desde) & " AND (((his_estructura.htethasta IS NULL OR his_estructura.htethasta <= " & ConvFecha(l_hasta) & ")) "
'                StrSql = StrSql & " OR (his_estructura.htethasta IS NULL OR his_estructura.htethasta >= " & ConvFecha(l_hasta) & "))) "
'                StrSql = StrSql & " INNER JOIN estr_cod ON his_estructura.estrnro = estr_cod.estrnro "
'                StrSql = StrSql & " INNER JOIN tipocod ON estr_cod.tcodnro = tipocod.tcodnro "
'                StrSql = StrSql & " WHERE Empleado.Ternro = " & rs("ternro")
'                StrSql = StrSql & " AND tipocod.tcodcodext = '155'" 'Modificada consulta ver 1.08

                StrSql = "SELECT DISTINCT his_estructura.estrnro, estr_cod.nrocod FROM empleado "
                StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.tenro = 4 " 'Puesto
                StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(l_hasta) & " AND ((his_estructura.htethasta IS NULL OR his_estructura.htethasta >= " & ConvFecha(l_desde) & ")) )"
                StrSql = StrSql & " INNER JOIN estr_cod ON his_estructura.estrnro = estr_cod.estrnro "
                StrSql = StrSql & " INNER JOIN tipocod ON estr_cod.tcodnro = tipocod.tcodnro "
                StrSql = StrSql & " WHERE Empleado.Ternro = " & rs("ternro")
                StrSql = StrSql & " AND tipocod.tcodcodext = '155'" 'Modificada consulta ver 1.08
                
                Flog.writeline "Puesto:" & StrSql
                OpenRecordset StrSql, rs2
                If rs2.EOF Then
                    Flog.writeline "No se encontró el Puesto para el empleado ternro(" & rs("ternro") & ")"
                    puesto = "00"
                Else
                    If rs2.RecordCount > 1 Then
                        'Agregado por Carmen Quintero
                       Flog.writeline "El empleado ternro(" & rs("ternro") & ") tiene mas de un puesto en el mes."
                       rs2.MoveLast
                       'Fin
                    End If
                    puesto = "0" & rs2("nrocod")
                End If
                rs2.Close
                '--------------------------Fin Puesto---------------------------------
                '-----------------------Licencias - - Agregado ver 1.03 --------------------
'                StrSql = " SELECT his_estructura.estrnro, estr_cod.nrocod FROM empleado "
'                StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.tenro = 30 " 'Situacion de Revista
'                StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(l_desde) & " AND (his_estructura.htethasta IS NULL OR his_estructura.htethasta >= " & ConvFecha(l_hasta) & "))"
'                StrSql = StrSql & " INNER JOIN estr_cod ON his_estructura.estrnro = estr_cod.estrnro "
'                StrSql = StrSql & " INNER JOIN tipocod ON estr_cod.tcodnro = tipocod.tcodnro "
'                StrSql = StrSql & " WHERE empleado.ternro = " & rs("ternro")
'                StrSql = StrSql & " AND tipocod.tcodcodext = '155' " 'Código para definir la licencia - Agregado ver 1.03
                
'                StrSql = "SELECT DISTINCT his_estructura.estrnro, estr_cod.nrocod FROM empleado "
'                StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.tenro = 30 " 'Situación de Revista
'                StrSql = StrSql & " AND (his_estructura.htetdesde >= " & ConvFecha(l_desde) & " AND (((his_estructura.htethasta IS NULL OR his_estructura.htethasta <= " & ConvFecha(l_hasta) & ")) "
'                StrSql = StrSql & " OR (his_estructura.htethasta IS NULL OR his_estructura.htethasta >= " & ConvFecha(l_hasta) & "))) "
'                StrSql = StrSql & " INNER JOIN estr_cod ON his_estructura.estrnro = estr_cod.estrnro "
'                StrSql = StrSql & " INNER JOIN tipocod ON estr_cod.tcodnro = tipocod.tcodnro "
'                StrSql = StrSql & " WHERE Empleado.Ternro = " & rs("ternro")
'                StrSql = StrSql & " AND tipocod.tcodcodext = '155'" 'Modificada consulta ver 1.08
                
                StrSql = "SELECT DISTINCT his_estructura.estrnro, estr_cod.nrocod FROM empleado "
                StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.tenro = 30 " 'Situación de Revista
                StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(l_hasta) & " AND ((his_estructura.htethasta IS NULL OR his_estructura.htethasta >= " & ConvFecha(l_desde) & "))) "
                StrSql = StrSql & " INNER JOIN estr_cod ON his_estructura.estrnro = estr_cod.estrnro "
                StrSql = StrSql & " INNER JOIN tipocod ON estr_cod.tcodnro = tipocod.tcodnro "
                StrSql = StrSql & " WHERE Empleado.Ternro = " & rs("ternro")
                StrSql = StrSql & " AND tipocod.tcodcodext = '155'" 'Modificada consulta ver 1.08
                OpenRecordset StrSql, rs2
                If rs2.EOF Then
                    Flog.writeline "No se encontró la Licencia para el empleado ternro(" & rs("ternro") & ")"
                    licencia = "00"
                Else
                    If rs2.RecordCount > 1 Then
                       'Agregado por Carmen Quintero
                       Flog.writeline "El empleado ternro(" & rs("ternro") & ") tiene mas de una situacion de revista en el mes."
                       rs2.MoveLast
                       'Fin
                    End If
                    If rs2("nrocod") < 10 Then
                        licencia = "0" & rs2("nrocod")
                    Else
                        licencia = rs2("nrocod")
                    End If
                End If
                rs2.Close
                
                '----------------------Fin Licencias----------------------------------------
                '------datos fijos: licencia y puesto en 00 y 05 respectivamente. Comentado ver 1.03
                'licencia = "00" Comentado ver 1.03
                'puesto = "04" Comentado ver 1.03
                
                
                
                'Gereracion de Archivo -----------------------------------------------------
                If CLng(hab_aportes) > 0 Then 'Si no tiene haberes con aporte, no imprime.
                'Al venir por confrep, lo filtro directamente en la impresión
                
                  Call imprimirTexto(cuil, ArchExp, 13, True)           'Cuil
                  ArchExp.Write Sep
                  Call imprimirTexto(Apellido, ArchExp, 30, True)       'apellido
                  ArchExp.Write Sep
                  Call imprimirTexto(Nombre, ArchExp, 30, True)        'nombre
                  ArchExp.Write Sep
                  Call prntTextoCeros(f_ingreso, ArchExp, 8, False)      'fecha de ingreso
                  ArchExp.Write Sep
                  Call imprimirTexto(med_jornada, ArchExp, 1, True)     'media jornada (s/n)
                  ArchExp.Write Sep
                  Call TextoScomaYnum(hab_aportes, ArchExp, 20, True)    'haberes con aportes
                  ArchExp.Write Sep
                  Call TextoScomaYnum(impor_noremun, ArchExp, 20, True)  'importe no remunerativo
                  ArchExp.Write Sep
                  Call imprimirTextoConCeros(categoria, ArchExp, 2, True)       'categoria
                  ArchExp.Write Sep
                  Call imprimirTexto(es_afiliado, ArchExp, 1, True)     'Es Afiliado
                  ArchExp.Write Sep
                  Call prntTextoCeros(f_baja, ArchExp, 8, False)          'fecha de baja
                  ArchExp.Write Sep
                  Call imprimirTexto(tip_docum, ArchExp, 1, True)       ' tipo de documento
                  ArchExp.Write Sep
                  Call imprimirTexto(nro_docum, ArchExp, 15, True)      'nummero de documento
                  ArchExp.Write Sep
                  Call imprimirTexto(puesto, ArchExp, 2, True)          ' Puesto
                  ArchExp.Write Sep
                  Call imprimirTexto(licencia, ArchExp, 2, True)        'Licencia
                  
                  
                  ArchExp.writeline
                End If 'del filtro por hab_aportes>0
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



