Attribute VB_Name = "ExpSecZonaN"
' ---------------------------------------------------------------------------------------------
' Descripcion: Genera exportacion SEC Zona Norte
' ---------------------------------------------------------------------------------------------
Option Explicit


Global Const Version = "1.01"
Global Const FechaModificacion = "02/05/2016"
Global Const UltimaModificacion = " " 'Se corrigen los parámetros en el caso de que vengan vacíos
'CAS-35073 - BDO - Nuevo formato de exportación SEC (CAS-15298) :Miriam Ruiz

'Global Const Version = "1.00"
'Global Const FechaModificacion = "26/01/2016"
'Global Const UltimaModificacion = " " 'Version Inicial - Se copia del SEC Rosario


'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
Global fs, f

Global NroProceso As Long

Global Path As String
Global HuboErrores As Boolean

Global EmpErrores As Boolean


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
Global Fechahasta
Global Fechadesde
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
    
    Nombre_Arch = PathFLog & "ExportacionSECZonaN" & "-" & NroProceso & ".log"
    
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
    StrSql = StrSql & " AND btprcnro = " & 461
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

Public Function Esta_de_Licencia_Tipo(ByVal Fechadesde As Date, ByVal Fechahasta As Date, ByVal Tercero As Long, ByVal Tipos As String) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna TRUE si el dia de la fecha de Licencia de algunos de los tipos está entre la fecha dese y hasta del período seleccionado . Sino FALSE.

' ---------------------------------------------------------------------------------------------
Dim rs_Lic As New ADODB.Recordset

    StrSql = "SELECT empleado FROM emp_lic WHERE (empleado = " & Tercero & " )"
    StrSql = StrSql & " AND elfechadesde <=" & ConvFecha(Fechahasta)
    StrSql = StrSql & " AND elfechahasta >= " & ConvFecha(Fechadesde)
    StrSql = StrSql & " AND tdnro IN (" & Tipos & ")"
    OpenRecordset StrSql, rs_Lic
    Esta_de_Licencia_Tipo = Not rs_Lic.EOF
    
    'cierro
    If rs_Lic.State = adStateOpen Then rs_Lic.Close
    Set rs_Lic = Nothing
End Function


Public Function ValidarRuta(ByVal modarchdefault As String, ByVal carpeta As String, ByVal msgError As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Valida que exista una carpeta, sino la crea y devuelve la ruta completa.
' Autor      : Gonzalez Nicolás
' Fecha      : 09/01/2012
' Ultima Mod : FGZ - 26/09/2012 - le agregué control por si la variable carpeta ya viene con barra "\"
' ---------------------------------------------------------------------------------------------
    'Activo el manejador de errores
    'On Error Resume Next
    Dim Ruta, carpetanueva
    Dim fs ' Licho - CAS-17249 - HEIDT- INSTALACION DEL SEC EN TEST80
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    'FGZ - 26/09/2012 --------------------------------
    If Left$(carpeta, 1) = "\" Or Left$(carpeta, 2) = "\\" Then
        Ruta = modarchdefault & carpeta
    Else
        Ruta = modarchdefault & "\" & carpeta
    End If
    'Ruta = modarchdefault & "\" & carpeta
    'FGZ - 26/09/2012 --------------------------------
    
    If Not fs.FolderExists(Ruta) Then
        Flog.writeline "La carpeta " & carpeta & " no existe. Se creará."
        'CREO LA CARPETA
        Set carpetanueva = fs.CreateFolder(Ruta)
        'Flog.writeline "error: " & Err.Number
        If Err.Number <> 0 Then
            'IMPRIME MSG. DE ERROR EN CASO QUE NO SE PUEDA CREAR LA CARPETA.
            Select Case msgError
                Case 1:
                    Flog.writeline "Error al crear la carpeta " & carpeta & " - Consulta al Administrador del Sistema."
            End Select
        End If
        ValidarRuta = Ruta
    Else
        ValidarRuta = Ruta
    End If
End Function


Private Sub Generar_Archivo(ByVal ArrParametros As Variant)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera la exportacion
' Autor      : Miriam Ruiz
' Fecha      : 25/01/2016

'CUIL                       11
'APELLIDO y NOMBRE          40
'FECHA INGRESO              10 (DD/MM/AAAA)
'HABERES C_APORTE           9
'CATEGORIA                  1 Fijo "1"
'Aportante                  1 Fijo "1"
'FECHA BAJA                 10  (DD/MM/AAAA)
'LICENCIA                   1 (0:ninguna, 1:Lic maternidad, 2: Lic S/goce de sueldo
'MEDIA JORNADA              1 (0= NO, 1=SI)

'SEPARADOR DE TODOS LOS CAMPOS: No usa
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





Dim Cuil As String '11 caracteres sin guiones.
Dim Apellido As String ' 40 caracteres, completa con blancos
Dim f_ingreso As String ' 10 digitos formato (DD/MM/AAAA)
Dim hab_aportes As String ' 9 dígitos, completa con ceros a derecha
Dim categoria As String ' 1 Fijo "1"
Dim Aportante As String ' 1 Fijo "1"
Dim f_baja As String ' 10 digitos formato (DD/MM/AAAA)
Dim licencia As String ' 1 dígitos, (0:ninguna, 1:Lic maternidad, 2: Lic S/goce de sueldo)
Dim med_jornada As String '1 dígito (0=no, 1=si)

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
Dim ListaLicmat As String
Dim ListaLicSGS As String
Dim Lista_SalNremun As String
Dim ListaMediajor As String
Dim ListaAfiliado As String


Lista_SalRemun = "0"
ListaLicmat = "0"
ListaLicSGS = "0"


Dim arrSindicatos

'Dim formatofecha As String

'Escribe log de ingreso
Flog.writeline "Ingresó a generar el archivo"


Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs_Aux As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim Ruta As String

    On Error GoTo ME_Local
    
    '////////////////////////////////////////////////////////////////////
 
        'l_filtro = CStr(ArrParametros(0))
         l_filtro = CStr(IIf(ArrParametros(0) = "null" Or ArrParametros(0) = "", "( 0=0 )", ArrParametros(0)))
        Flog.writeline "Filtro: " & l_filtro
     
        'fecha estructura
      
        l_fecestr = CDate(IIf(ArrParametros(1) = "null" Or ArrParametros(1) = "", Date, ArrParametros(1)))
        'Estructura 1
       ' l_tenro1 = CInt(ArrParametros(2))
       'l_estrnro1 = CInt(ArrParametros(3))
        l_tenro1 = CInt(IIf(ArrParametros(2) = "null" Or ArrParametros(2) = "", 0, ArrParametros(2)))
        l_estrnro1 = CInt(IIf(ArrParametros(3) = "null" Or ArrParametros(3) = "", 0, ArrParametros(3)))
        'Estructura 2
        'l_tenro2 = CInt(ArrParametros(4))
        'l_estrnro2 = CInt(ArrParametros(5))
         l_tenro2 = CInt(IIf(ArrParametros(4) = "null" Or ArrParametros(4) = "", 0, ArrParametros(4)))
        l_estrnro2 = CInt(IIf(ArrParametros(5) = "null" Or ArrParametros(5) = "", 0, ArrParametros(5)))
        'Estructura 3
        'l_tenro3 = CInt(ArrParametros(6))
        'l_estrnro3 = CInt(ArrParametros(7))
        l_tenro3 = CInt(IIf(ArrParametros(6) = "null" Or ArrParametros(6) = "", 0, ArrParametros(6)))
        l_estrnro3 = CInt(IIf(ArrParametros(7) = "null" Or ArrParametros(7) = "", 0, ArrParametros(7)))
        'Periodo liquidacion desde
        l_pliqdesde = CLng(IIf(ArrParametros(8) = "null" Or ArrParametros(8) = "", 0, ArrParametros(8)))
        'Periodo liquidacion hasta
        l_pliqhasta = CLng(IIf(ArrParametros(9) = "null" Or ArrParametros(9) = "", 0, ArrParametros(9)))
        'fecha periodo desde
        'l_desde = C_Date(ArrParametros(10))
        l_desde = CDate(IIf(ArrParametros(10) = "null" Or ArrParametros(10) = "", Date, ArrParametros(10)))
        'fecha periodo hasta
        l_hasta = CDate(IIf(ArrParametros(11) = "null" Or ArrParametros(11) = "", Date, ArrParametros(11)))
        'Procesos estados
        '   Aprob.Definitivo = 3
        '   Aprob.Provisorio = 2
        '   Liquidado = 1
        '   No Liquidado = 0
        '   Todos = -1
        l_proaprob = CLng(IIf(ArrParametros(12) = "", 0, ArrParametros(12))) 'CInt(ArrParametros(12))
        ' Lista de Procesos
        'l_listaproc = CStr(ArrParametros(13))
         l_listaproc = CStr(IIf(ArrParametros(13) = "null" Or ArrParametros(13) = "", "0", ArrParametros(13)))
        Flog.writeline "Procesos seleccionados: " & l_listaproc
        ' Conseptos - no se usa
        'l_concnro = CLng(ArrParametros(14))
        'Empresa nro
        'l_empresa = CLng(ArrParametros(15))
        l_empresa = CLng(IIf(ArrParametros(15) = "null" Or ArrParametros(15) = "", 0, ArrParametros(15)))
        'Nombre concepto no se usa
        'l_conceptonombre = CStr(ArrParametros(16))
        'Orden
        'l_orden = CStr(ArrParametros(17))
         l_orden = CStr(IIf(ArrParametros(17) = "null" Or ArrParametros(17) = "", "0", ArrParametros(17)))
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
    NroModelo = 2010 '
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
      Flog.writeline "Encontró el número de modelo"
        If Not IsNull(rs_Modelo!modarchdefault) Then
         Ruta = ValidarRuta(directorio, rs_Modelo!modarchdefault, 0)
            'directorio = directorio & Trim(rs_Modelo!modarchdefault)
            directorio = Ruta
            If Right(directorio, 1) <> "\" Then
                directorio = directorio & "\"
            End If
        Else
            Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta destino. El archivo será generado en el directorio default"
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo " & NroModelo & ". El archivo será generado en el directorio default"
    End If

    'Obtengo los datos del separador
    Sep = rs_Modelo!modseparador
    SepDec = rs_Modelo!modsepdec
    Flog.writeline Espacios(Tabulador * 0) & "Separador seleccionado: " & Sep
    
    'Seteo el nombre del archivo a exportar en el directorio indicado
   
    Nombre_Arch = directorio & rs_Modelo!modarchdefault & "-" & NroProceso & ".txt"
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
    NroReporte = 504
    
  
    Lista_SalRemun = "0"
    ListaLicmat = "0"
    ListaLicSGS = "0"
    ListaMediajor = "0"
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
                Case "AC"
                      Lista_SalRemun = Lista_SalRemun & "," & rs!confval
                Case "EST"
                      ListaMediajor = ListaMediajor & "," & rs!confval
                Case "LMA"
                      ListaLicmat = ListaLicmat & "," & rs!confval
                Case "SGS"
                      ListaLicSGS = ListaLicSGS & "," & rs!confval
            End Select
            rs.MoveNext
        Loop
    End If
    
    
    'Pone a funcionar, de los filtros seleccionados, los que se eligen en ordenamiento
    If l_tenro3 <> 0 Then ' esto ocurre solo cuando se seleccionan los tres niveles

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
                
                
                '-Acumula Haberes con aportes ----------------
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
            
                
                
                '**************************************************************
          
                'Busco el cuil del empleado
                StrSql = " SELECT nrodoc "
                StrSql = StrSql & " FROM ter_doc "
                StrSql = StrSql & " WHERE ternro = " & rs("ternro")
                StrSql = StrSql & " AND tidnro = 10 "
                OpenRecordset StrSql, rs2
                If rs2.EOF Then
                    Flog.writeline "No se encontró el Cuil para el empleado.ternro(" & rs("ternro") & ")"
                    Cuil = "00000000000"
                Else
                    Cuil = Left(CStr(rs2("nrodoc")), 13)
                    Cuil = Replace(CStr(Cuil), "-", "")
                    Cuil = Right("00000000" & Cuil, 11)
                    Cuil = Left(CStr(Cuil), 11)
                End If
                rs2.Close
                
              
                'Busco la fecha de alta y baja
           
                f_ingreso = "00/00/0000"
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
                        f_ingreso = "00/00/0000"
                        f_baja = "00/00/0000"
                        rs2.Close
                    Else
                        If Not EsNulo(rs2("empfaltagr")) Then
                            f_ingreso = CStr(CDate(rs2("empfaltagr")))
                           Else
                           f_ingreso = "00/00/0000"
                        End If
                        'La fecha de baja puede, y es común, que no exista.
                        If Not EsNulo(rs2("empfecbaja")) Then
                          f_baja = CStr(CDate(rs2("empfecbaja")))
                        Else
                          f_baja = "00/00/0000"
                        End If
                      
                        rs2.Close
                    End If
                Else
                    f_ingreso = CStr(CDate(rs2("altfec")))
                    If Not EsNulo(rs2("bajfec")) Then
                      f_baja = CStr(CDate(rs2("bajfec")))
                    Else
                      f_baja = "00/00/0000"
                    End If
                   
                    rs2.Close
                End If
                'rs2.Close
                
                              
                'traigo los datos de nombre y apellido
                StrSql = "SELECT ternom,ternom2, terape, terape2 "
                StrSql = StrSql & "FROM empleado "
                StrSql = StrSql & "WHERE empleado.ternro = " & rs("ternro")
                OpenRecordset StrSql, rs2
                         
                 
                'verifico que tenga datos
                If rs2.EOF Then
                     Apellido = ""
                Else
                     Apellido = rs2("terape") & " " & rs2("terape2") & " " & rs2("ternom") & " " & rs2("ternom2")
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
                
            
                '-----------------------Licencias -  --------------------
                licencia = "0"
                If Esta_de_Licencia_Tipo(l_desde, l_hasta, rs("ternro"), ListaLicmat) Then
                    licencia = "1"
                End If
                If Esta_de_Licencia_Tipo(l_desde, l_hasta, rs("ternro"), ListaLicSGS) Then
                    licencia = "2"
                End If
                '----------------------Fin Licencias----------------------------------------

                
                
                
                'Gereracion de Archivo -----------------------------------------------------
                If CLng(hab_aportes) > 0 Then 'Si no tiene haberes con aporte, no imprime.
                'Al venir por confrep, lo filtro directamente en la impresión
                 
                  Call imprimirTexto(Cuil, ArchExp, 11, True)           'Cuil
                  Call imprimirTexto(Apellido, ArchExp, 40, True)       'apellido
                  Call imprimirTexto(f_ingreso, ArchExp, 10, False)     'fecha de ingreso
                  Call TextoScomaYnum(hab_aportes, ArchExp, 9, True)    'haberes con aportes
                  Call imprimirTexto("1", ArchExp, 1, True)             'categoria
                  Call imprimirTexto("1", ArchExp, 1, True)             'Aportante
                  Call imprimirTexto(f_baja, ArchExp, 10, False)        'fecha de baja
                  Call imprimirTexto(licencia, ArchExp, 1, True)        'Licencia
                  Call imprimirTexto(med_jornada, ArchExp, 1, True)     'media jornada (s/n)
                  
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



