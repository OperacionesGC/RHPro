Attribute VB_Name = "ExpUOM"
' ---------------------------------------------------------------------------------------------
' Descripcion: Proyecto encargado de generar la exportacion Interface UOM
' Autor      : FGZ
' Fecha      : 17/01/2006
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaModificacion = "17/01/2006"
'Global Const UltimaModificacion = " " 'Version Inicial

'Global Const Version = "1.02"
'Global Const FechaModificacion = "17/02/2006"
'Martin Ferraro - 17/02/2006 - Se agrego "distinct" en la consulta principal de empleados
'y se agrego funcion CeroIzq para acomodar los decimales despuesd del punto de la remuneracion

'Global Const Version = "1.03"
'Global Const FechaModificacion = "20/12/2006"
'Global Const UltimaModificacion = " " 'FGZ - 20/12/2006
'                                       Le cambié terestrciv x estcivnro

'Global Const Version = "1.04"
'Global Const FechaModificacion = "15/03/2007"
'Global Const UltimaModificacion = " " 'LM - 20/12/2006
''                                       Se cambio el orden entre localidad y codigo postal
''                                       Se modifico la seleccion de sindicato
''                                       Se cambio el formato de la fecha de exportacion a DDMMAAAA

'Global Const Version = "1.05"
'Global Const FechaModificacion = "17/08/2007"
'Global Const UltimaModificacion = " " 'FGZ - Se cambio que filtre empleados de los convenios configurados y no de los sindicatos
'                                       Configuracion en Confrep 155

'Global Const Version = "1.06"
'Global Const FechaModificacion = "02/03/2010"
'Global Const UltimaModificacion = " " 'EGO - Se comentaron datos q no van mas en el informe y se modifico formato de otros.

'Global Const Version = "1.07"
'Global Const FechaModificacion = "11/04/2011"
'Global Const UltimaModificacion = " Habilitación de campos no obligatorios "
'Verónica Bogado - Se descomentaron campos no obligatorios para que aparezcan en el archivo.


'Global Const Version = "1.08"
'Global Const FechaModificacion = "02/06/2011"
'Global Const UltimaModificacion = " Validación de tipo de Documento y situación de Revista para Oracle"

'Global Const Version = "1.09"
'Global Const FechaModificacion = "01/08/2011"
'Global Const UltimaModificacion = " Error bof or eof en consulta de situacion de revista."

'Global Const Version = "1.10"
'Global Const FechaModificacion = "18/11/2013"
'Global Const UltimaModificacion = " Se modifico la consulta para que busque el cod. ext de la categoria."
'Borrelli Facundo -  Se modifico la consulta para que busque el cod. ext de la categoria

'Global Const Version = "1.11"
'Global Const FechaModificacion = "22/11/2013"
'Global Const UltimaModificacion = " Se modifico la consulta para que muestre el cod. ext. de la provincia de la sucursal."
'Borrelli Facundo -  Se modifico la consulta para que muestre el cod. ext. de la provincia de la sucursal

Global Const Version = "1.12"
Global Const FechaModificacion = "28/11/2013"
Global Const UltimaModificacion = "Error en la consulta, tenia la fecha hasta fija, cuando busca el cod. ext. de la provincia de la sucursal."
'Borrelli Facundo -  CAS 22338 - Error en la consulta, tenia la fecha hasta fija, cuando busca el cod. ext. de la provincia de la sucursal."

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
Global fs, f
'Global flog
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
Global concnro As Integer
Global Conccod As String
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
' Autor      : FGZ
' Fecha      : 17/01/2006
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim rs As New ADODB.Recordset
Dim PID As String
Dim Parametros As String
Dim ArrParametros

Dim Sucursal As Long
Dim Periodo As Long


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
    
    NroProceso = NroProcesoBatch
    
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
    
    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ExportacionUOM" & "-" & NroProceso & ".log"
    
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
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 0) & "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND btprcnro = 121"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
       'Obtengo los parametros del proceso
       Parametros = rs!bprcparam
       ArrParametros = Split(Parametros, "@")
       
       Sucursal = CLng(ArrParametros(0))
       Periodo = CLng(ArrParametros(1))
      
       Call Generar_Archivo(Sucursal, Periodo)
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
    If u < 0 Then 'el texto es más grande que la longitud especificada
        cadena = Left(cadena, Longitud)
        archivo.Write cadena
    ElseIf u > 0 Then
        If derecha Then
            archivo.Write cadena & String(u, " ")
        Else
            archivo.Write String(u, " ") & cadena
        End If
    Else 'u es cero, el texto tiene la longitud esperada
      archivo.Write cadena
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

Private Sub Generar_Archivo(ByVal Sucursal As Long, ByVal Periodo As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera la exportacion
' Autor      : FGZ
' Fecha      : 17/01/2006
' Ultima Mod : 26/02/2010 Elizabeth Gisela Oviedo
' Descripcion:

' Listado dia 26/02/2010
'Orden Campo Posición Longitud
'1 cuil 1 11
'2 nombre 12 30
'3 situacion 42 3
'4 sindicato 45 1
'5 sueldo 46 8

'Listado anterior
''Tipo de Documento:         1   3 Caracteres
''Nrodoc:                    4   8
''Cuil:                      12  11
''Nombre:                    23  30
''Domicilio Calle:           53  20
''Domicilio Nro:             73  5
''Domicilio Piso:            78  3
''Domicilio Dpto:            81  4
''Domicilio Cod Postal:      85  7
''Domicilio Localidad        92  20
''Telefono                   112 15
''Estado CIvil:              127 3
''Sexo:                      130 1
''Nacionalidad:              131 3
''Fecha Nacimiento:          134 8
''Fecha Alta:                142 8
''Fecha Baja:                150 8
''Situacion:                 158 3
''Sindicato:                 161 1
''Remuneracion Total:        162 10
''Cod Osocial                172 6


' ---------------------------------------------------------------------------------------------
'Dim objRs As New ADODB.Recordset
Dim Nombre_Arch As String
Dim NroModelo As Long
Dim Directorio As String
Dim Carpeta
Dim fs1

Dim Sep As String
Dim SepDec As String

Dim cantRegistros As Long
Dim Empresa As String

Dim Legajo As String
Dim Tercero As Long

Dim TDoc As String
Dim Nrodoc As String
Dim Cuil As String
Dim nombre As String
Dim Calle As String
Dim Nro As String
Dim Piso As String
Dim Dpto As String
Dim CodPostal As String
Dim Localidad As String
Dim Telefono As String
Dim Provincia As String
Dim EstadoCivil As String
Dim Sexo As String
Dim Nacionalidad As String
Dim FechaNacimiento As String
Dim FechaAlta As String
Dim FechaBaja As String
Dim Situacion As String
Dim Sindicato As String
Dim Remuneracion As Double
Dim RemunAux As String
Dim CodOSocial As String
Dim NroReporte As Long
Dim Acumulador As Long
Dim Incapacidad
'Dim Sindicato1 As Long
'Dim Sindicato2 As Long
Dim Categoria As String
Dim Lista_Sindicatos As String
Dim Lista_Convenios As String
Dim FhastaConv
Dim FdesdeConv

Dim arrSindicatos
Dim I As Integer
Dim formatofecha As String
formatofecha = "YYYYMMDD"
Dim CodigoSit As Integer

Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs_codSit As New ADODB.Recordset
Dim rs_Aux As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
    
    On Error GoTo ME_Local
      
    NroModelo = 273
    'Directorio de exportacion
    StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Directorio = Trim(rs!sis_dirsalidas)
    End If
    
    
    StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
    OpenRecordset StrSql, rs_Modelo
    If Not rs_Modelo.EOF Then
        If Not IsNull(rs_Modelo!modarchdefault) Then
            Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
        Else
            Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta desteino. El archivo será generado en el directorio default"
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo " & NroModelo & ". El archivo será generado en el directorio default"
    End If
    
    '-----------------------------//--------------------------------------
    'Directorio temporal para debug
    'Directorio = "C:\log"
             
    'Obtengo los datos del separador
    Sep = ";" 'rs_Modelo!modseparador Predefinido en ;
    SepDec = rs_Modelo!modsepdec
    Flog.writeline Espacios(Tabulador * 0) & "Separador seleccionado: " & Sep
     
    Nombre_Arch = Directorio & "\ExpUom" & "-" & NroProceso & ".txt"
    Flog.writeline Espacios(Tabulador * 1) & "Se crea el archivo: " & Nombre_Arch
    Set fs = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
        Set Carpeta = fs1.CreateFolder(Directorio)
    End If
    'desactivo el manejador de errores
    Set ArchExp = fs.CreateTextFile(Nombre_Arch, True)


    On Error GoTo ME_Local
    
    'Configuracion del Reporte
    NroReporte = 155
    StrSql = "SELECT * FROM confrep "
    StrSql = StrSql & " WHERE repnro = " & NroReporte
    StrSql = StrSql & " AND confnrocol = 1"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Flog.writeline "No se encontró la configuración del Reporte"
        Exit Sub
    Else
        Acumulador = rs!confval
    End If
    
    
    'Oscar esto es lo que hay que hacer, igualmente probalo
    '**************************
    Lista_Sindicatos = "0"
    StrSql = "SELECT * FROM confrep "
    StrSql = StrSql & " WHERE repnro = " & NroReporte
    'FGZ - 17/08/2007 - se modificó esta linea
    'StrSql = StrSql & " AND confnrocol > 1"
    StrSql = StrSql & " AND confnrocol > 1 AND confnrocol <= 10 "
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Flog.writeline "No hay sindicatos configurados"
        Exit Sub
    End If
    Do While Not rs.EOF
        Lista_Sindicatos = Lista_Sindicatos & "," & rs!confval
        rs.MoveNext
    Loop
    
    
    '**************************
    
    'FGZ - 17/08/2007 - Se agregó esto -----------------
    Lista_Convenios = "0"
    StrSql = "SELECT * FROM confrep "
    StrSql = StrSql & " WHERE repnro = " & NroReporte
    StrSql = StrSql & " AND confnrocol > 10"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Flog.writeline "No hay Convenios Configurados. No Saldrá Ningun Empleado."
        Exit Sub
    End If
    Do While Not rs.EOF
        Lista_Convenios = Lista_Convenios & "," & rs!confval
        rs.MoveNext
    Loop
    'FGZ - 17/08/2007 - Se agregó esto -----------------
        
    'Cargo las fechas desde y Hasta
    StrSql = "SELECT * FROM periodo WHERE pliqnro = " & Periodo
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
     
    If Not rs.EOF Then
       FechaDesde = ConvFecha(rs!pliqdesde)
       FechaHasta = ConvFecha(rs!pliqhasta)
       FhastaConv = rs!pliqhasta
       FdesdeConv = rs!pliqdesde
    Else
       Flog.writeline "No se encontro el periodo " & Periodo
       Exit Sub
    End If
    
    
    
    '------------------------------------------------------------------
    'Busco los datos
    '------------------------------------------------------------------
    StrSql = " SELECT distinct empleado.empleg, empleado.ternro, sindicato.estrnro as codsindictato "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " INNER JOIN his_estructura sucursal ON empleado.ternro = sucursal.ternro AND sucursal.tenro = 1 AND sucursal.estrnro = " & Sucursal
    StrSql = StrSql & " INNER JOIN his_estructura sindicato ON empleado.ternro = sindicato.ternro AND sindicato.tenro = 16 "
    StrSql = StrSql & " INNER JOIN his_estructura convenio ON empleado.ternro = convenio.ternro AND convenio.tenro = 19 "
    StrSql = StrSql & " AND ( "
    'Sindicato1 = 653
    'Sindicato2 = 649
'   Esta linea no va, es de prueba
    'Sindicato2 = 1428
    'StrSql = StrSql & " sindicato.estrnro = " & Sindicato1
    'StrSql = StrSql & " OR sindicato.estrnro = " & Sindicato2
    
    '****************************
    'Oscar esto es lo que hay que hacer, igualmente probalo
    'FGZ - 17/08/2007 - se cambió esto
    'StrSql = StrSql & " sindicato.estrnro IN (" & Lista_Sindicatos & ")"
    StrSql = StrSql & " convenio.estrnro IN (" & Lista_Convenios & ")"
    '****************************
    
    StrSql = StrSql & ")"
    StrSql = StrSql & " WHERE "
    'Sucursal
    StrSql = StrSql & " sucursal.htetdesde <= " & FechaHasta
    StrSql = StrSql & " AND (sucursal.htethasta IS NULL OR sucursal.htethasta >= " & FechaDesde & ")"
    'Sindicato
    StrSql = StrSql & " AND sindicato.htetdesde <= " & FechaHasta
    StrSql = StrSql & " AND (sindicato.htethasta IS NULL OR sindicato.htethasta >= " & FechaDesde & ")"
    'Convenio
    StrSql = StrSql & " AND convenio.htetdesde <= " & FechaHasta
    StrSql = StrSql & " AND (convenio.htethasta IS NULL OR convenio.htethasta >= " & FechaDesde & ")"
    
    StrSql = StrSql & " ORDER BY empleado.empleg "
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    
    
    'seteo de las variables de progreso
    Progreso = 0
    If Not rs.EOF Then
        cantRegistros = rs.RecordCount
        If cantRegistros = 0 Then
           cantRegistros = 1
           Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos a Exportar."
        End If
    Else
        cantRegistros = 1
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos a Exportar."
    End If
    IncPorc = (99 / cantRegistros)
    
    Do While Not rs.EOF
        Tercero = rs!Ternro
        Legajo = rs!empleg
        Flog.writeline "Legajo: " & Legajo
       
       
        'FGZ - 13/02/2006
        'Si el empleado no tiene liquidaciones en el mes ==> no sale en la exportacion
        
        StrSql = "SELECT cabliq.cliqnro FROM cabliq "
        StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
        StrSql = StrSql & " INNER JOIN periodo ON proceso.pliqnro = periodo.pliqnro "
        StrSql = StrSql & " WHERE periodo.pliqnro = " & Periodo
        StrSql = StrSql & " AND cabliq.empleado = " & Tercero
        OpenRecordset StrSql, rs_Procesos
        If Not rs_Procesos.EOF Then
             ' ----------------------------------------------------------------
            'Sindicato
            Sindicato = "N"
            arrSindicatos = Split(Lista_Sindicatos, ",")
            For I = 0 To UBound(arrSindicatos)
                If CLng(rs!codsindictato) = CLng(arrSindicatos(I)) Then
                    Sindicato = "S"
                End If
            Next
       
            ' ----------------------------------------------------------------
            ' Busco Tipo y nro de doc
            Flog.writeline "Busco Tipo y nro de doc"
            StrSql = " SELECT ter_doc.nrodoc, tipodocu.tidcod_bco FROM tercero "
            StrSql = StrSql & " INNER JOIN ter_doc ON tercero.ternro = ter_doc.ternro AND ter_doc.tidnro <= 4 "
            StrSql = StrSql & " INNER JOIN tipodocu ON tipodocu.tidnro = ter_doc.tidnro "
            StrSql = StrSql & " WHERE tercero.ternro= " & Tercero
            If rs_Aux.State = adStateOpen Then rs_Aux.Close
            OpenRecordset StrSql, rs_Aux
            If Not rs_Aux.EOF Then
              If EsNulo(rs_Aux!tidcod_bco) Then
                TDoc = ""
              Else
                TDoc = Left(CStr(rs_Aux!tidcod_bco), 3)
              End If
                Nrodoc = Left(CStr(rs_Aux!Nrodoc), 8)
            Else
                Flog.writeline "Error al obtener los datos del cuil"
                TDoc = "   "
                Nrodoc = "00000000"
            End If
         
            ' ----------------------------------------------------------------
            ' Buscar el CUIL
            Flog.writeline "Buscar el CUIT"
            StrSql = " SELECT cuil.nrodoc FROM tercero "
            StrSql = StrSql & " INNER JOIN ter_doc cuil ON (tercero.ternro = cuil.ternro AND cuil.tidnro = 10) "
            StrSql = StrSql & " WHERE tercero.ternro= " & Tercero
            If rs_Aux.State = adStateOpen Then rs_Aux.Close
            OpenRecordset StrSql, rs_Aux
            If Not rs_Aux.EOF Then
                Cuil = Left(CStr(rs_Aux!Nrodoc), 13)
                Cuil = Replace(CStr(Cuil), "-", "")
                Cuil = Replace(CStr(Cuil), ".", "")
                Cuil = Left(CStr(Cuil), 11)
            Else
                Flog.writeline "Error al obtener los datos del cuil"
                Cuil = "00000000000"
            End If
           
            ' ----------------------------------------------------------------
            'Buscar el apellido y nombre
            Flog.writeline "Buscar el apellido y nombre"
            StrSql = "SELECT * FROM tercero WHERE ternro = " & Tercero
            If rs_Aux.State = adStateOpen Then rs_Aux.Close
            OpenRecordset StrSql, rs_Aux
            If Not rs_Aux.EOF Then
                nombre = Left(rs_Aux!terape & " " & rs_Aux!ternom, 30)
                
                'Sexo
                If rs_Aux!tersex Then
                    Sexo = "M"
                Else
                    Sexo = "F"
                End If
                
                'Nacionalidad
                If Not EsNulo(rs_Aux!nacionalnro) Then
                    StrSql = " SELECT nacionalcodext FROM nacionalidad WHERE nacionalnro = " & rs_Aux!nacionalnro
                    If rs2.State = adStateOpen Then rs2.Close
                    OpenRecordset StrSql, rs2
                    If Not rs2.EOF Then
                        Nacionalidad = Left(IIf(Not EsNulo(rs2!nacionalcodext), rs2!nacionalcodext, "   "), 3)
                    Else
                        Nacionalidad = "" 'Space(3)
                        Flog.writeline "No se encontró la nacionalidad."
                    End If
                Else
                    Nacionalidad = "" 'Space(3)
                    Flog.writeline "No se encontró la nacionalidad."
                End If
                
                'Fecha de Nacimiento
                If Not EsNulo(rs_Aux!terfecnac) Then
                    FechaNacimiento = Format(rs_Aux!terfecnac, formatofecha)
                Else
                    FechaNacimiento = ""
                    Flog.writeline "No se encontró la fecha de nacimiento."
                End If
            Else
                Flog.writeline "No se encontró el tercero."
                Exit Sub
            End If

                
                'FGZ - 20/12/2006
                'Cambié terestciv x estcivnro
                'Estado civil
                If Not EsNulo(rs_Aux!estcivnro) Then
                    StrSql = " SELECT extciv_bco FROM estcivil WHERE estcivnro = " & rs_Aux!estcivnro
                    If rs2.State = adStateOpen Then rs2.Close
                    OpenRecordset StrSql, rs2
                    If Not rs2.EOF Then
                        If Not EsNulo(rs2!extciv_bco) Then
                           EstadoCivil = Left(rs2!extciv_bco, 3)
                        Else
                            EstadoCivil = "" 'Space(3)
                            Flog.writeline "El estado civil no tiene configurado el codigo bco."
                        End If
                    Else
                        EstadoCivil = "" 'Space(3)
                        Flog.writeline "No se encontró el estado civil."
                    End If
                Else
                    StrSql = " SELECT extciv_bco FROM estcivil WHERE estcivnro = " & rs_Aux!estcivnro
                    If rs2.State = adStateOpen Then rs2.Close
                    OpenRecordset StrSql, rs2
                    If Not rs2.EOF Then
                        EstadoCivil = Left(rs2!extciv_bco, 3)
                    Else
                        EstadoCivil = "" 'Space(3)
                        Flog.writeline "No se encontró el estado civil."
                    End If
                End If
                
                                
                'Busco datos de la provincia del empleado
                StrSql = "SELECT provincia.provnro, provcodext FROM his_estructura"
                StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = his_estructura.ternro"
                StrSql = StrSql & " INNER JOIN sucursal ON his_estructura.estrnro=sucursal.estrnro"
                StrSql = StrSql & " INNER JOIN cabdom ON sucursal.ternro=cabdom.ternro"
                StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
                StrSql = StrSql & " INNER JOIN provincia ON provincia.provnro = detdom.provnro"
                StrSql = StrSql & " Where his_estructura.tenro = 1 And Empleado.Ternro = " & Tercero
                'StrSql = StrSql & " AND htetdesde <=" & FechaDesde & " AND"
                'StrSql = StrSql & " (htethasta >=" & FechaHasta & " OR htethasta IS NULL)"
                'FB - 22/11/2013 - Se modifico la consulta para que muestre el cod. ext. de la provincia de la sucursal
                StrSql = StrSql & " AND ((his_estructura.htetdesde <= " & FechaDesde & " ) "
                StrSql = StrSql & " OR ( his_estructura.htetdesde >= " & FechaDesde & " )) "
                'FB - 28/11/2013 - Se corrigio la consulta donde quedaba fija la fecha hasta.
                'StrSql = StrSql & " AND ((his_estructura.htethasta >= 31/11/2013 ) "
                StrSql = StrSql & " AND ((his_estructura.htethasta >= " & FechaHasta & " )"
                StrSql = StrSql & " or (his_estructura.htethasta is null)) "
                OpenRecordset StrSql, rs_Aux
                
                If Not rs_Aux.EOF Then
                  Provincia = rs_Aux!provcodext
                Else
                  Flog.writeline "No se pudo establecer la provincia de la sucursal del empleado. Verificar configuración"
                  Provincia = "" 'Space(2)
                End If
                
                
                
           
           
            ' ----------------------------------------------------------------
            'Buscar el domicilio
            Flog.writeline "Buscar el domicilio"
            StrSql = " SELECT * FROM detdom "
            StrSql = StrSql & " INNER JOIN zona ON zona.zonanro = detdom.zonanro "
            StrSql = StrSql & " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro "
            StrSql = StrSql & " INNER JOIN localidad ON localidad.locnro = detdom.locnro "
            StrSql = StrSql & " WHERE cabdom.ternro = " & Tercero
            StrSql = StrSql & " and cabdom.tidonro=2"
            If rs_Aux.State = adStateOpen Then rs_Aux.Close
            OpenRecordset StrSql, rs_Aux
            If Not rs_Aux.EOF Then
                Calle = Left(CStr(IIf(Not IsNull(rs_Aux!Calle), rs_Aux!Calle, "")), 20)
                Nro = Left(CStr(IIf(Not IsNull(rs_Aux!Nro), rs_Aux!Nro, "")), 5)
                Piso = Left(CStr(IIf(Not IsNull(rs_Aux!Piso), rs_Aux!Piso, "")), 3)
                Dpto = Left(CStr(IIf(Not IsNull(rs_Aux!oficdepto), rs_Aux!oficdepto, "")), 4)
                CodPostal = Left(CStr(IIf(Not IsNull(rs_Aux!codigopostal), rs_Aux!codigopostal, "")), 7)
                Localidad = Left(CStr(IIf(Not IsNull(rs_Aux!locdesc), rs_Aux!locdesc, "")), 20)
            Else
                Flog.writeline "No se encontraron datos del domicilio tipo 2 (Particular). SQL : " & StrSql
                Calle = "" 'Space(20)
                Nro = "" 'Space(5)
                Piso = "" 'Space(3)
                Dpto = "" 'Space(4)
                CodPostal = "" 'Space(7)
                Localidad = "" 'Space(20)
            End If
           
    
            ' ----------------------------------------------------------------
            'Buscar el Telefono
            Flog.writeline "Buscar el Telefono"
            StrSql = " SELECT * FROM telefono "
            StrSql = StrSql & " INNER JOIN cabdom ON telefono.domnro = cabdom.domnro "
            StrSql = StrSql & " WHERE cabdom.ternro = " & Tercero
            StrSql = StrSql & " AND telefono.teldefault = -1"
            If rs_Aux.State = adStateOpen Then rs_Aux.Close
            OpenRecordset StrSql, rs_Aux
            If Not rs_Aux.EOF Then
                Telefono = Left(CStr(IIf(Not IsNull(rs_Aux!telnro), rs_Aux!telnro, "")), 15)
            Else
                Flog.writeline "No se encontraron datos del telefono. SQL : " & StrSql
                Telefono = "" 'Space(15)
            End If
           
           
            'Fecha de alta y fecha de baja
'            StrSql = "SELECT * FROM fases WHERE real = -1 AND empleado = " & Tercero
'            StrSql = StrSql & " AND altfec <= " & ConvFecha(FechaHasta)
'            StrSql = StrSql & " ORDER BY altfec"
'            If rs_Aux.State = adStateOpen Then rs_Aux.Close
'            OpenRecordset StrSql, rs_Aux
'            If Not rs_Aux.EOF Then
'                rs_Aux.MoveLast
'                FechaAlta = Format(rs_Aux!altfec, formatofecha)
'                If Not EsNulo(rs_Aux!bajfec) Then
'                    FechaBaja = Format(rs_Aux!bajfec, formatofecha)
'                Else
'                    FechaBaja = Space(8)
'                End If
'            Else
'                Flog.writeline "No se encontraron fases reales. SQL : " & StrSql
'                FechaAlta = Space(8)
'                FechaBaja = Space(8)
'            End If
       Flog.writeline " Busco el tipo de codigo para la situación revista UOM (Columna 20)"
       StrSql = " SELECT confval tipocodsit FROM confrep "
       StrSql = StrSql & " WHERE repnro = " & NroReporte
       StrSql = StrSql & " and confnrocol = 20"
           OpenRecordset StrSql, rs_codSit
             If Not rs_codSit.EOF Then
              CodigoSit = CInt(rs_codSit!tipocodsit)
             Else
             CodigoSit = 0
             Flog.writeline " No encontro tipo de código Situación de Revista configurado en la columna 20 "
             End If
            
            If rs_codSit.State = adStateOpen Then
            rs_codSit.Close
            End If
            
            'Situacion de revista
            ' ----------------------------------------------------------------
            Flog.writeline "Buscar Estructura Situacion de Revista Actual"
            StrSql = " SELECT estructura.estrnro FROM estructura "
            StrSql = StrSql & " INNER JOIN his_estructura ON estructura.estrnro = his_estructura.estrnro "
            StrSql = StrSql & " WHERE his_estructura.ternro = " & Tercero & " AND "
            StrSql = StrSql & " his_estructura.tenro = 30 AND "
            StrSql = StrSql & " ((his_estructura.htetdesde <= " & FechaDesde & " ) or ( his_estructura.htetdesde >= " & FechaDesde & " )) AND "
            StrSql = StrSql & " ((his_estructura.htethasta >= " & FechaHasta & " ) or (his_estructura.htethasta is null))"
            StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
            Flog.writeline StrSql
            If rs_Aux.State = adStateOpen Then
            rs_Aux.Close
            End If
            OpenRecordset StrSql, rs_Aux
            'If Not rs_Aux.EOF And Not rs_Aux!estrcodext = "" Then
            If Not rs_Aux.EOF Then
               'If EsNulo(rs_Aux!Estrnro) Then
                ' Flog.writeline " No se encontro estructura  de la Situacion de Revista."
                'Situacion = ""
                'Else
                'Situacion = Left(rs_Aux!estrcodext, 3)
                
                  Flog.writeline " Se encontro estructura de la Situación de Revista: " & rs_Aux!Estrnro
                
                  StrSql = " SELECT nrocod codigosituacion  FROM estr_cod "
                  StrSql = StrSql & " WHERE tcodnro = " & CodigoSit
                  StrSql = StrSql & " AND  estrnro = " & rs_Aux!Estrnro
                   OpenRecordset StrSql, rs_codSit
                    If Not rs_codSit.EOF Then
                 
                     Situacion = Left(rs_codSit!codigosituacion, 3)
                  
                     Flog.writeline " Codigo Externo de la situacion encontrado " & Situacion
                    Else
                     Flog.writeline " No se encontro el código de Situación de revista configurada en la estructura  " & rs_Aux!Estrnro & "Tipo codigo ha asociar:" & CodigoSit
                   End If
                   rs_codSit.Close
                    
                'End If
            Else
                Flog.writeline " No se encontro encontro estructura de revista para periodo " & FechaDesde & " - " & FechaHasta & " . SQL : " & StrSql
                Situacion = "" ' space 3
          End If
           
            'Remuneracion Total
            StrSql = "SELECT ammonto FROM acu_mes "
            StrSql = StrSql & " WHERE acunro = " & Acumulador
            StrSql = StrSql & " AND ternro = " & Tercero
            StrSql = StrSql & " AND ammes = " & Month(FhastaConv)
            StrSql = StrSql & " AND amanio = " & Year(FhastaConv)
            If rs_Aux.State = adStateOpen Then rs_Aux.Close
            OpenRecordset StrSql, rs_Aux
            If Not rs_Aux.EOF Then
                If EsNulo(rs_Aux!ammonto) Then
                    Remuneracion = 0
                Else
                    Remuneracion = Format(rs_Aux!ammonto, "#####0.00")
                End If
            Else
                Flog.writeline "No se encontró remuneración. SQL : " & StrSql
                Remuneracion = 0
            End If
            
            ' ----------------------------------------------------------------
            ' Buscar la Obra Social del empleado
'            Flog.writeline "Buscar la Obra Social del empleado"
'            StrSql = " SELECT * FROM his_estructura "
'            StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
'            StrSql = StrSql & " WHERE his_estructura.ternro = " & Tercero & " AND "
'            StrSql = StrSql & " his_estructura.tenro = 17 AND "
'            StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(FechaHasta) & ") AND "
'            StrSql = StrSql & " ((" & ConvFecha(FechaHasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
'            StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
'            If rs_Aux.State = adStateOpen Then rs_Aux.Close
'            OpenRecordset StrSql, rs_Aux
'            If Not rs_Aux.EOF Then
'                StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Aux!Estrnro
'                StrSql = StrSql & " AND tcodnro = 1"
'                If rs_Aux.State = adStateOpen Then rs_Aux.Close
'                OpenRecordset StrSql, rs_Aux
'                If Not rs_Aux.EOF Then
'                    CodOSocial = Left(rs_Aux!nrocod, 6)
'                Else
'                    Flog.writeline "No se encontró el codigo interno para la Obra Social"
'                    CodOSocial = "000000"
'                End If
'            Else
'                Flog.writeline "No se encontro la Obra Social"
'                CodOSocial = "000000"
'            End If

            'Busca condición de incapacidad
            StrSql = "select empdiscap from empleado where ternro=" & Tercero
            OpenRecordset StrSql, rs_Aux
            Flog.writeline "Sql incapacidad: " & StrSql
            If Not rs_Aux.EOF Then
              If rs_Aux!empdiscap = 0 Then
                Incapacidad = "N"
              Else
                Incapacidad = "S"
              End If
            Else
              Flog.writeline "No se encontraron datos que especifiquen si el empleado tiene una incapacidad"
              Incapacidad = "" 'Space(1)
            End If
            rs_Aux.Close
                              
            'Busco la categoria
            StrSql = "select estrcodext from estructura inner join his_estructura "
            StrSql = StrSql & "on estructura.estrnro=his_estructura.estrnro"
            StrSql = StrSql & " where estructura.tenro=3 and his_estructura.ternro=" & Tercero
            'StrSql = StrSql & " and htetdesde<= " & FechaDesde & ""
            'StrSql = StrSql & " and (htethasta>= " & FechaHasta & " or htethasta is null)"
            'FB - 18/11/2013 - Se modifico la consulta para que busque el cod. ext de la categoria
            'Se agregaron las siguientes lineas para que tome de forma correcta la fecha desde y hasta
            StrSql = StrSql & " AND ((his_estructura.htetdesde <= " & FechaDesde & " ) or ( his_estructura.htetdesde >= " & FechaDesde & " )) AND "
            StrSql = StrSql & " ((his_estructura.htethasta >= " & FechaHasta & " ) or (his_estructura.htethasta is null))"
            'Hasta aca
            OpenRecordset StrSql, rs_Aux
            If Not rs_Aux.EOF Then
              If EsNulo(rs_Aux!estrcodext) Then
              Flog.writeline "No se encontró Codigo externo de la Categoría."
              Else
              Categoria = rs_Aux!estrcodext
              Flog.writeline " Codigo Externo de la Categoria Encontrado " & Categoria
              End If
            Else
              Flog.writeline "No se encontró información de Categoría"
              Categoria = "" 'Space(2)
            End If
            rs_Aux.Close
            
            'Extraigo el nro de documento del cuil (por si no es dni)
            Nrodoc = Right(Cuil, 9)
            Nrodoc = Left(Nrodoc, 8)
                     
            'escribo en el archivo
'            Call imprimirTexto(TDoc, ArchExp, 3, True)          'Tipo de doc
            Call imprimirTextoNro(Cuil, ArchExp, 11, True)         'CUIL
            Call imprimirTextoNro(nombre, ArchExp, 30, True)       'Nombre
            'Call imprimirTexto(Nrodoc, ArchExp, 8, True)        'Nro de documento, queda dentro del cuil
            Call imprimirTextoNro(Situacion, ArchExp, 2, True)      'Situacion de revista
            Call imprimirTextoNro(Sindicato, ArchExp, 1, True)      'Sindicato
            If Remuneracion = 0 Then
                Call imprimirTextoNro("0,00", ArchExp, 8, True)  'Remuneracion Total
            Else
                RemunAux = CerozIzq(Remuneracion)
                RemunAux = Replace(CStr(RemunAux), ".", ",")
                Call imprimirTextoNro(RemunAux, ArchExp, 8, True)  'Remuneracion Total
            End If
            Call imprimirTextoNro(Nrodoc, ArchExp, 8, True)         'Nro Documento
            Call imprimirTextoNro(Calle, ArchExp, 32, True)        'calle
            Call imprimirTextoNro(Nro, ArchExp, 5, True)           'Nro
            'Call imprimirTexto(Piso, ArchExp, 3, True)          'Piso
            'Call imprimirTexto(Dpto, ArchExp, 4, True)          'Dpto
            Call imprimirTextoNro(Localidad, ArchExp, 13, True)    'Localidad
            Call imprimirTextoNro(CodPostal, ArchExp, 4, True)     'Codigo postal
            Call imprimirTextoNro(Provincia, ArchExp, 2, True)
            'Call imprimirTexto(Telefono, ArchExp, 15, True)      'Telefono
            'FGZ - 13/02/2006 Alinear a derecha
'            Call imprimirTexto(Telefono, ArchExp, 15, False)      'Telefono'
            Call imprimirTextoNro(EstadoCivil, ArchExp, 2, True)    'Estado civil
            Call imprimirTextoNro(Sexo, ArchExp, 1, True)           'sexo
            Call imprimirTextoNro(Nacionalidad, ArchExp, 3, True)   'Nacionalidad
            Call imprimirTextoNro(FechaNacimiento, ArchExp, 8, True) 'Fecha de nacimiento
'            Call imprimirTexto(FechaAlta, ArchExp, 8, True)      'Fecha de alta
'            Call imprimirTexto(FechaBaja, ArchExp, 8, True)      'Fecha de baja
            
'            Call imprimirTexto(CodOSocial, ArchExp, 6, True)     'Codigo de OSocial <--------
            Call imprimirTextoNro(Incapacidad, ArchExp, 1, True)
            Call imprimirTextoNro(Categoria, ArchExp, 2, True)
            ArchExp.writeline
            
            'Actualizo el estado del proceso
            TiempoAcumulado = GetTickCount
            cantRegistros = cantRegistros - 1
            Progreso = Progreso + IncPorc
            
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
            StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
            StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        Else
            Flog.writeline "No se encontraron liquidaciones para el legajo en el periodo"
        End If
        
        rs.MoveNext
    Loop
    ArchExp.Close
    
    
    'Cierro y libero todo
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs_Aux.State = adStateOpen Then rs_Aux.Close
    Set rs_Aux = Nothing
    If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
    Set rs_Modelo = Nothing
    If rs2.State = adStateOpen Then rs2.Close
    Set rs2 = Nothing
    If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
    Set rs_Procesos = Nothing

Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
End Sub



