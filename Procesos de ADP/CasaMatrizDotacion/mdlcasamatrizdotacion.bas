Attribute VB_Name = "mdlcasamatrizDotacion"
' __________________________________________________________________________________________________
' Descripcion: Reporte Dotación Santillana
' Autor      : Margiotta, Emanuel
' Fecha      : 24/02/2010
' Ultima Mod :
' Descripcion:
' ___________________________________________________________________________________________________

Option Explicit


'Global Const Version = "1.00"
'Global Const FechaModificacion = " 23-02-2010 "
'Global Const UltimaModificacion = " " 'Margiotta, Emanuel

'Global Const Version = "1.01"
'Global Const FechaModificacion = " 29-03-2010 "
'Global Const UltimaModificacion = " " 'FGZ
''                                       Para los totales se consideran solo los meses generados hasta la fecha del filtro
''                                       Se redondea a nro entero el total de empleados

'Global Const Version = "1.02"
'Global Const FechaModificacion = " 05-04-2010"
'Global Const UltimaModificacion = " " 'Margiotta Emanuel
''                                       Se Corrigieron fechas de las estructuras

'Global Const Version = "1.03"
'Global Const FechaModificacion = " 28-04-2010"
'Global Const UltimaModificacion = " " 'FGZ
''                                       Se Corrigieron calculo de bajas

Global Const Version = "1.04"
Global Const FechaModificacion = " 08-10-2010"
Global Const UltimaModificacion = " " 'Margiotta, Emanuel
'                                       Se cambio el tipo de datos a Long de algunas variables porque daba desbordamiento

' ________________________________________________________________________________________

Global NroProceso As Long
Global Path As String
Global HuboErrores As Boolean


Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset


Global filtro As String  ' filtro trae si el empleado es activo o no, y legajo desde -hasta
Global tenro1 As Integer
Global estrnro1 As Integer
Global tenro2 As Integer
Global estrnro2 As Integer
Global tenro3 As Integer
Global estrnro3 As Integer
Global agencia As Integer
Global fecestr As Date
Global repnro As Integer
Global empleados As String
Global Empresas As String
Global contratos_eventuales As String
Global MesCorte As Integer

Dim cargo As Integer
Dim Anio As Integer

'EAM
Dim arrEstructura() As String
Dim arrTipoEstructura() As String   'Tiene los tipo de estructura asociado a cada Nivel
Dim cantElementos As Integer
Dim Coordenadas() As Integer
Dim NombreNiveles() As String       'Array de 2 Dimensiones. Tiene los nombres  de los niveles
Dim NomNivel(5) As String           'Array de 1 Dimension, Busca en NombreNiveles() y lo guarda en este para el insert
Dim NombreSeccion(4) As String

Dim AC As String
Dim CausaDesp As String

Dim IdUser As String
Dim bpfecha As Date
Dim bphora As String

' Global fecestr As String
'Global TituloRep As String
Global HayDetalleLiq(1 To 12) As Double


Private Sub Main()
 Dim strCmdLine As String
 Dim Nombre_Arch As String

 Dim StrSql As String

 Dim PID As String
 Dim Parametros As String
 Dim ArrParametros
 Dim cantRegistros
 

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
    
    'Carga las configuraciones basicas, formato de fecha, string de conexion,
    'tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas


    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "RepDotacionCasaMatriz" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)


    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    Flog.writeline
    
    Flog.writeline "Inicio Proceso de Reporte Dotación a Casa Matriz : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 0) & "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        'Obtengo los parametros del proceso
        IdUser = rs!IdUser
        bpfecha = rs!bprcfecha
        bphora = rs!bprchora
        Parametros = rs!bprcparam
        
        ArrParametros = Split(Parametros, "@")
             
        'Carga los parametros del proceso en var Globales
        Call levantarParametros(ArrParametros)
                  
        'Carga las estructuras asociadas a cada nivel, configuradas en el reporte
        Call CargarConfiguracionReporte
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 "
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & 0 & "'"
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'Borrar los datos por si se reprocesa - Detalle
        StrSql = "DELETE FROM rep_dot_cm_det " & _
                 "WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'Cabecera
        StrSql = "DELETE FROM rep_dot_cm_cab " & _
                 "WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
                        
        ' _____________________________________________________
        ' Armar consulta Ppal según Filtro - empleado con estructuras activas a la Fecha
        ' ____________________________________________________
        
        Call filtro_empleados(StrSql, fecestr)
        OpenRecordset StrSql, rs1
       
        'Seteo de las variables de progreso
        Progreso = 5
        Call actualizar_progreso(Progreso)
                        
        'Comienza el Proceso
        Call InsertarDatosCab
    
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No se encontraron los datos del proceso Nro " & NroProceso
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
    
    'Flog.writeline "cant open " & Cantidad_de_OpenRecordset

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

' ___________________________________________________________________________________________________
' EAM-  Inserta la cabecera del reporte.
' ---------------------------------------------------------------------------------------------------
Sub InsertarDatosCab()
  Dim StrSql As String

On Error GoTo MError

    Flog.writeline Espacios(Tabulador * 1) & "Insertar datos de la cabecera  "

    StrSql = "INSERT INTO rep_dot_cm_cab " & _
             "(bpronro,Fecha,Hora,tenro1,estrnro1,tenro2, estrnro2, tenro3, estrnro3, fecestr,anio) " & _
             "VALUES (" & NroProceso & "," & ConvFecha(bpfecha) & ",'" & bphora & "'," & tenro1 & "," & estrnro1 & "," & _
             tenro2 & "," & estrnro2 & "," & tenro3 & "," & estrnro3 & "," & ConvFecha(fecestr) & "," & Anio & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

    'Inserta el detalle del reporte
    Call InsertarDatosdet

Exit Sub

MError:
    HuboErrores = True
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQL: " & StrSql
End Sub


'EAM- Calcula la cantidad de elementos que tiene la matriz
Private Function ElementosMatriz() As Integer
 Dim Nivel, I, j, cant, elem As Integer
       
    elem = 0
    cant = 1
    Nivel = UBound(arrEstructura, 1)
    For I = 1 To Nivel
        For j = 1 To UBound(arrEstructura, 2)
            If arrEstructura(I, j) <> "" Then
                elem = elem + 1
            End If
        Next
        cant = cant * elem
        elem = 0
    Next
    ElementosMatriz = cant
End Function

'Inicializa el arreglo de Cordenadas porque la ultima posicion debe tener un 0
'Ej. 3 niveles --> (1,1,0)
Private Sub InicializarArrCordenadas(IndCoordenadas() As Integer, Nivel As Integer)
 Dim I As Integer
    ReDim IndCoordenadas(5)
    
    For I = 1 To Nivel - 1
        IndCoordenadas(I) = 1
    Next
    IndCoordenadas(I) = 0
        
End Sub

'EAM- Calcula todas las convinaciones entre los niveles que tiene la matriz
'Devuelve un arregeglo con los indices siguientes de cada nivel para evaluar
Private Sub SigConjuntoIndice(Coordenadas() As Integer, Optional Inicializar As Boolean)
 Dim Nivel As Integer
                          
    Nivel = UBound(arrEstructura, 1)
    
    If Inicializar Then
        Call InicializarArrCordenadas(Coordenadas(), Nivel)
    End If
    
        
    While (Nivel > 0)
    
        If UBound(arrEstructura, 2) > Coordenadas(Nivel) Then
            If arrEstructura(Nivel, Coordenadas(Nivel) + 1) <> "" Then
                Coordenadas(Nivel) = Coordenadas(Nivel) + 1
                Nivel = 0
            Else
                Coordenadas(Nivel) = 1
            End If
        Else
            Coordenadas(Nivel) = 1
        End If
                
        Nivel = Nivel - 1
    Wend
    
End Sub

'EAM- Calcula los totales y subtotales
Private Sub CalcularTotalSubTotal(ByVal Seccion As Integer, ByRef TotalSeccion() As Double, ByVal CalcularTotal As Boolean, SubSec1 As Integer, SubSec2 As Integer, SubSec3 As Integer, Optional subsec1desc As String, Optional subsec2desc As String, Optional subsec3desc As String)
 Dim rsdet As New ADODB.Recordset
 
    'Obtiene la cantidad de estructuras del nivel 1 para calcular subtotal
    StrSql = "SELECT DISTINCT confval FROM confrep WHERE confrep.repnro= 273 AND conftipo= 'EN1'"
    
    'Seteo el total en 0. se calcula abajo
    'TotalSeccion(13) = 0
        
    OpenRecordset StrSql, rs
    While Not rs.EOF
        StrSql = "SELECT sum(m1) m1, sum(m2) m2, sum(m3) m3, sum(m4) m4, sum(m5) m5, sum(m6) m6, sum(m7) m7, sum(m8) m8, sum(m9) m9, " & _
                 "sum(m10) m10, sum(m11) m11, sum(m12) m12, sum(total) total " & _
                 "FROM rep_dot_cm_det where bpronro=" & NroProceso & " AND n1= " & rs!confval & " AND sec= " & Seccion & _
                 " AND subsec1= " & SubSec1 & " AND subsec2= " & SubSec2 & " AND subsec3= " & SubSec3
        OpenRecordset StrSql, rsdet
        
        'Inserta los subtotales
        StrSql = "Insert INTO rep_dot_cm_det (bpronro,sec,subsec1,subsec2,subsec3, secdesc,subsec1desc,subsec2desc,subsec3desc, " & _
                 "n1,n2,n3,n4,n5,n1desc,n2desc,n3desc,n4desc,n5desc," & _
                 "m1,m2,m3,m4,m5,m6,m7,m8,m9,m10,m11,m12,total) VALUES (" & NroProceso & "," & Seccion & "," & SubSec1 & "," & _
                 SubSec2 & "," & SubSec3 & ",'" & NombreSeccion(Seccion) & "','" & subsec1desc & "','" & subsec2desc & "','" & _
                 subsec3desc & "'," & rs!confval & "," & _
                 0 & "," & 0 & "," & 0 & "," & 0 & ",'" & "" & "','" & "" & "','" & _
                 "" & "','" & "" & "','" & "" & "'," & _
                 IIf(Not EsNulo(rsdet!m1), rsdet!m1, 0) & "," & IIf(Not EsNulo(rsdet!m2), rsdet!m2, 0) & "," & IIf(Not EsNulo(rsdet!m3), rsdet!m3, 0) & "," & IIf(Not EsNulo(rsdet!m4), rsdet!m4, 0) & "," & IIf(Not EsNulo(rsdet!m5), rsdet!m5, 0) & "," & IIf(Not EsNulo(rsdet!m6), rsdet!m6, 0) & "," & _
                 IIf(Not EsNulo(rsdet!m7), rsdet!m7, 0) & "," & IIf(Not EsNulo(rsdet!m8), rsdet!m8, 0) & "," & IIf(Not EsNulo(rsdet!m9), rsdet!m9, 0) & "," & IIf(Not EsNulo(rsdet!m10), rsdet!m10, 0) & "," & _
                 IIf(Not EsNulo(rsdet!m11), rsdet!m11, 0) & "," & IIf(Not EsNulo(rsdet!m12), rsdet!m12, 0) & "," & rsdet!total & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        TotalSeccion(13) = TotalSeccion(13) + rsdet!total
        rs.MoveNext
    Wend

    'Calcula el total de la Seccion
    If CalcularTotal Then
        'Inserta el total de la Seccion
        StrSql = "Insert INTO rep_dot_cm_det (bpronro,sec,subsec1,subsec2, subsec3, secdesc,subsec1desc,subsec2desc,subsec3desc,n1,n2,n3,n4,n5,n1desc,n2desc,n3desc,n4desc,n5desc," & _
                 "m1,m2,m3,m4,m5,m6,m7,m8,m9,m10,m11,m12,total) VALUES (" & NroProceso & "," & Seccion & "," & SubSec1 & "," & _
                 SubSec2 & "," & SubSec3 & ",'" & NombreSeccion(Seccion) & "','" & subsec1desc & "','" & subsec2desc & "','" & _
                 subsec3desc & "'," & 0 & ","
        StrSql = StrSql & 0 & "," & 0 & "," & 0 & "," & 0 & ",'" & "" & "','" & "" & "','" & "" & "','" & "" & "','" & "" & "',"
        StrSql = StrSql & TotalSeccion(1) & "," & TotalSeccion(2) & "," & TotalSeccion(3) & "," & TotalSeccion(4) & "," & TotalSeccion(5) & "," & TotalSeccion(6)
        StrSql = StrSql & "," & TotalSeccion(7) & "," & TotalSeccion(8) & "," & TotalSeccion(9) & "," & TotalSeccion(10)
        StrSql = StrSql & "," & TotalSeccion(11) & "," & TotalSeccion(12) & "," & TotalSeccion(13) & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End Sub


'EAM- Carga los parametros del proceso
Sub levantarParametros(ArrParametros)

On Error GoTo ME_param
    
    'Numero Desde-Hasta de leg. y el estado
    filtro = Replace(ArrParametros(0), "..", "'", 1, Len(ArrParametros(0)))
    'filtro = ArrParametros(0)
    tenro1 = CInt(ArrParametros(1))
    estrnro1 = CInt(ArrParametros(2))
    tenro2 = CInt(ArrParametros(3))
    estrnro2 = CInt(ArrParametros(4))
    tenro3 = CInt(ArrParametros(5))
    estrnro3 = CInt(ArrParametros(6))
    agencia = CInt(ArrParametros(7))
    fecestr = CStr(ArrParametros(8))    'Fecha de la estructuras
    cargo = CInt(ArrParametros(9))
    'repnro = ArrParametros(10)
    Anio = CInt(ArrParametros(11))
    Empresas = ArrParametros(12)
    MesCorte = Month(fecestr)
    
    Flog.writeline Espacios(Tabulador * 0) & "PARAMETROS"
    Flog.writeline Espacios(Tabulador * 0) & "Filtro: " & filtro
    Flog.writeline Espacios(Tabulador * 0) & "Tenro1: " & tenro1
    Flog.writeline Espacios(Tabulador * 0) & "Estrnro1: " & estrnro1
    Flog.writeline Espacios(Tabulador * 0) & "Tenro2: " & tenro2
    Flog.writeline Espacios(Tabulador * 0) & "Estrnro2: " & estrnro2
    Flog.writeline Espacios(Tabulador * 0) & "Tenro3: " & tenro3
    Flog.writeline Espacios(Tabulador * 0) & "Estrnro3: " & estrnro3
    Flog.writeline Espacios(Tabulador * 0) & "Agencia: " & agencia
    Flog.writeline Espacios(Tabulador * 0) & "Fecha p/Estruct: " & fecestr
    Flog.writeline Espacios(Tabulador * 0) & "Cargo: " & cargo
    Flog.writeline Espacios(Tabulador * 0) & "Nro Reporte: " & repnro
    Flog.writeline Espacios(Tabulador * 0) & "Año: " & Anio
    Flog.writeline Espacios(Tabulador * 0) & "Empresas: " & Empresas

Exit Sub

ME_param:
    Flog.writeline "    Error: Error en la carga de Parametros "
    
End Sub

'Carga la configuración del confRep
'cantElementos -> Tiene la cantidad de convinaciones de la matriz de estructuras
Sub CargarConfiguracionReporte()
 Dim rs As New ADODB.Recordset
 Dim rs1 As New ADODB.Recordset
 Dim rs2 As New ADODB.Recordset
 Dim strEstructura As String  'Concatena todas las estructuras del nivel
 Dim I, j, Y As Integer
 Dim l_tenro
 
 On Error GoTo ME_conf
        
    Flog.writeline Espacios(Tabulador * 1) & "Buscar la configuracion del Reporte - confrep  "
   
    'Obtiene el maximo elemento Configurado para armar el arreglo
    StrSql = "SELECT * FROM confrep WHERE confrep.repnro= 273 AND conftipo like 'EN%'  ORDER BY confval DESC"
    OpenRecordset StrSql, rs
    
    Y = rs!confval
     
    'Obtiene la cantidad de niveles
    StrSql = "SELECT * FROM confrep WHERE confrep.repnro= 273 AND conftipo= 'N'" & _
             " ORDER BY confnrocol DESC"
    OpenRecordset StrSql, rs
    
    ReDim arrEstructura(0 To rs!confval, 0 To Y)
    ReDim arrTipoEstructura(rs!confval)
    ReDim NombreNiveles(0 To rs!confval, 0 To Y)
        
    'Guarda los Tipos de Estructura de cada Nivel
    While Not rs.EOF
        arrTipoEstructura(rs!confval) = rs!confval2
        rs.MoveNext
    Wend
    rs.MoveFirst
   
       
    If rs.EOF Then
        Flog.writeline Espacios(Tabulador * 0) & "Error. Se debe configurar el confrep. Nro de confrep:" & repnro
        Exit Sub
    Else
        
        For I = 1 To UBound(arrEstructura, 1)
            'Obtiene la catidad de estructuras de cada nivel
            StrSql = "SELECT * FROM confrep WHERE confrep.repnro= 273 AND conftipo= 'EN" & I & "'" & _
                     " ORDER BY confnrocol DESC"
            OpenRecordset StrSql, rs1
            
            strEstructura = 0
        
            For j = 1 To rs1!confval
                
                'Obtiene las estructuras de cada nivel y orden EN1 - 4
                StrSql = "SELECT * FROM confrep WHERE confrep.repnro= 273 AND conftipo= 'EN" & I & "' AND confval=" & j & " " & _
                         " ORDER BY confnrocol ASC"
                OpenRecordset StrSql, rs2
                
                
                'Obtiene el nombre del nivel
                NombreNiveles(I, j) = rs2!confetiq
                
                If arrTipoEstructura(I) = -1 Then
                    arrEstructura(I, j) = rs2!confval2
                    strEstructura = strEstructura & "," & rs2!confval2
                Else
                
                If (Not rs2.EOF) And (rs2!confval2 <> -1) Then
                    arrEstructura(I, j) = rs2!confval2
                    strEstructura = strEstructura & "," & rs2!confval2
         
                Else
                    If rs2!confval2 = -1 Then
              
                        '== -1 obtiene todas las Estructuras menos las que tiene configuradas arriba de ella
                        StrSql = "SELECT * FROM confrep WHERE confrep.repnro= 273 AND conftipo= 'N' AND confnrocol=" & I
                        OpenRecordset StrSql, rs
                        l_tenro = rs!confval2
              
                        StrSql = "SELECT * FROM estructura WHERE tenro=" & l_tenro & " and not estrnro in (select estrnro from estructura where estrnro in (" & strEstructura & "))"
                        OpenRecordset StrSql, rs
                        arrEstructura(I, j) = rs!estrnro
                        rs.MoveNext
              
                        Do While Not rs.EOF
                            arrEstructura(I, j) = arrEstructura(I, j) & "," & rs!estrnro
                            rs.MoveNext
                        Loop
                    End If
                End If
                End If  'Fin ArrEstructura()
                
                'Si el Nivel tiene mas de una 1 estructura confg. entra aca
                rs2.MoveNext
                Do While Not rs2.EOF
                    arrEstructura(I, j) = arrEstructura(I, j) & "," & rs2!confval2
                    strEstructura = strEstructura & "," & rs2!confval2
                    rs2.MoveNext
                Loop
       
            Next
        Next
    End If

    'Obtiene la cantidad de convinaciones de la matriz
    cantElementos = ElementosMatriz
    
    'Carga los acumuladores configurados en el rep
    StrSql = "SELECT confetiq FROM confrep  WHERE Conftipo='SEC' AND repnro= 273 ORDER BY confval"
    OpenRecordset StrSql, rs
    For I = 1 To rs.RecordCount
        NombreSeccion(I) = rs!confetiq
        rs.MoveNext
    Next
    rs.Close
    
    'Carga los acumuladores configurados en el rep
    StrSql = "SELECT confval2 FROM confrep  WHERE Conftipo='AC' AND repnro= 273"
    OpenRecordset StrSql, rs
    AC = 0
    While Not rs.EOF
        AC = AC & "," & rs!confval2
        rs.MoveNext
    Wend
    
    'Carga las Causas de despido Configuradas en el confrep
    StrSql = "SELECT confval2 FROM confrep  WHERE Conftipo='CD' AND repnro= 273"
    OpenRecordset StrSql, rs
    CausaDesp = 0
    While Not rs.EOF
        CausaDesp = CausaDesp & "," & rs!confval2
        rs.MoveNext
    Wend
    
Exit Sub
ME_conf:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub

'EAM - Inserta el detalle de las distintas secciones
Sub InsertarDatosdet()
 Dim Secciones As Integer
 Dim I As Integer

On Error GoTo MError
    
    'Carga los acumuladores configurados en el rep
    StrSql = "SELECT Distinct confetiq FROM confrep  WHERE Conftipo='SEC' AND repnro= 273"
    OpenRecordset StrSql, rs
    Secciones = 4
    
    For I = 1 To rs.RecordCount
        Flog.writeline " "
        Flog.writeline Espacios(Tabulador * 1) & "Procesando el detalle de la Sección" & I & " - " & NombreSeccion(I)
    
        Select Case I
            Case 1:
                Call CalcularPM(I)
            Case 2:
                Call CalcularPlantillaActiva(I)
            Case 3:
                Call CalcularSeccionTres(I)
            Case 4:
                Call CalcularRotacion(I)
        End Select
    Next
  
Exit Sub

MError:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub

'Calcula la Plantilla Media - Calculo proporcional de horas y días
Sub CalcularPM(Seccion As Integer)
 Dim rsdet As New ADODB.Recordset
 Dim FechaDesde As Date
 Dim FechaHasta As Date
 Dim StrSql As String
 Dim I, j As Integer
 Dim NroMes As Integer
 Dim EstructuraNivel As String
 Dim nroCabRep As Long
 Dim porcentaje As Double
 Dim total As Double
 Dim TotalSeccion(13) As Double
 Dim SqlPorcentajeMes As String
 Dim horas As Integer
 Dim dias As Integer
 Dim logEstructura As String
 Dim SqlSexo As String
          
    
    Flog.writeline Espacios(Tabulador * 1) & "Sección " & Seccion & "Equivalencia Hombre - Mes (Plantilla Media)"
    
    'Inicializa variables
    Call SigConjuntoIndice(Coordenadas, True)
    total = 0
    IncPorc = (22.5 / cantElementos)
    
    
    Flog.writeline " "
    Flog.writeline Espacios(Tabulador * 1) & "Comienzo de Inserción de datos de la seccción " & Seccion
        
    'Recorre tantas veces como convinaciones haya (cantElementos)
    For I = 0 To cantElementos - 1
        
        'Recorre por los mes para calcular el porcentaje de cada mes del año
        For NroMes = 1 To MesCorte
            FechaHasta = ultimo_dia_mes(NroMes, Anio)
            FechaDesde = primer_dia_mes(NroMes, Anio)
        
            'Obtiene todos los empleados de acurdo al filtro y a las estructuras configuradas en el Rep.
            StrSql = "SELECT DISTINCT empleado.empleg, empleado.ternro, ammonto dias, e2.estrcodext horas " & _
                     "FROM empleado " & _
                     "INNER JOIN tercero ON empleado.ternro= tercero.ternro " & _
                     "INNER JOIN fases ON empleado.ternro = fases.empleado " & _
                     "INNER JOIN his_estructura as Emp ON Emp.ternro = empleado.ternro AND Emp.tenro = 10 "
                               
            'Filtra la empresa si se selecciono 1 sino no la filtra
            If Empresas <> -1 Then
                StrSql = StrSql & " AND Emp.estrnro = " & Empresas
                StrSql = StrSql & " AND Emp.htetdesde <= " & ConvFecha(FechaHasta) & " AND ((Emp.htethasta >=" & ConvFecha(FechaHasta) & ") OR Emp.htethasta is null)"
            End If
                    
            'Arma el Inner y el IN de la consulta para los conjuntos de Estructuras ConfRep
            For j = 1 To UBound(arrTipoEstructura)
                If arrTipoEstructura(j) <> -1 Then
                    StrSql = StrSql & " INNER JOIN his_estructura as he" & j & " ON he" & j & ".ternro = empleado.ternro AND he" & j & ".tenro = " & arrTipoEstructura(j) & _
                                      " AND he" & j & ".htetdesde <= " & ConvFecha(FechaHasta) & " AND ((he" & j & ".htethasta >=" & ConvFecha(FechaHasta) & ") OR he" & j & ".htethasta is null)"
                    EstructuraNivel = EstructuraNivel & "AND he" & j & ".estrnro IN ( " & arrEstructura(j, Coordenadas(j)) & ") "
                    logEstructura = logEstructura & arrEstructura(j, Coordenadas(j))
                    NomNivel(j) = NombreNiveles(j, Coordenadas(j))
                Else
                    NomNivel(j) = NombreNiveles(j, Coordenadas(j))
                    SqlSexo = "AND tersex= " & arrEstructura(j, Coordenadas(j))
                End If
            Next

            StrSql = StrSql & " LEFT JOIN his_estructura as he6 ON he6.ternro = empleado.ternro AND he6.tenro = 21 " & _
                              " AND he6.htetdesde <= " & ConvFecha(FechaHasta) & " AND ((he6.htethasta >=" & ConvFecha(FechaHasta) & ") OR he6.htethasta is null)" & _
                              " INNER JOIN estructura e2 ON he6.estrnro = e2.estrnro" & _
                              " INNER JOIN acu_mes ON acu_mes.ternro = empleado.ternro" & _
                              " WHERE acunro IN (" & AC & ") " & _
                              " AND ammes = " & Month(FechaHasta) & _
                              " AND amanio = " & Year(FechaHasta) & " "
            StrSql = StrSql & EstructuraNivel
            StrSql = StrSql & " AND empleado.ternro IN " & empleados
            StrSql = StrSql & " AND (fases.altfec <= " & ConvFecha(FechaHasta) & " AND ( fases.bajfec >= " & ConvFecha(FechaHasta) & " OR fases.bajfec IS NULL))"
            StrSql = StrSql & SqlSexo
            
            OpenRecordset StrSql, rsdet
            
            
            EstructuraNivel = ""
                        
            If Not rsdet.EOF Then
                'Flog.writeline Espacios(Tabulador * 2) & " Se encontraron. Estructuras: " & rsdet.RecordCount & " en la estructuras empleados " & logEstructura & " a la fecha" & FechaDesde & " - " & FechaHasta
            Else
                Flog.writeline Espacios(Tabulador * 2) & " No se encontraron empleados. Estructuras: " & logEstructura & " a la fecha " & FechaDesde & " - " & FechaHasta
            End If
            Flog.writeline
                        
            
            'Recorre todos los empleados encontrados
            While Not rsdet.EOF
                Flog.writeline Espacios(Tabulador * 4) & " empleado: " & rsdet!empleg
                
                'reviso la cantidad de hs diarias segun el regimen horario
                If EsNulo(rsdet!horas) Then
                    horas = 0
                    Flog.writeline Espacios(Tabulador * 5) & " No se configuro el regimen horario del empleado:" & rsdet("ternro")
                Else
                    horas = rsdet!horas
                End If
                Flog.writeline Espacios(Tabulador * 4) & " horas diarias: " & horas
            
                'Topeo la cantidad de dias
                If rsdet!dias > 30 Then
                    dias = 30
                Else
                    dias = rsdet!dias
                End If
                
                Flog.writeline Espacios(Tabulador * 4) & " dias: " & dias
                porcentaje = porcentaje + Round((1 * (dias) / 30) * (horas / 8), 2)
        
                rsdet.MoveNext
            Wend
            
            total = total + Round(porcentaje)
            TotalSeccion(NroMes) = TotalSeccion(NroMes) + Round(porcentaje)
            
            If NroMes = 1 Then
                'Flog.writeline Espacios(Tabulador * 1) & "Calculo el porcentaje de horas: " & porcentaje
                'Flog.writeline Espacios(Tabulador * 1) & "Sql: " & StrSql
                
                StrSql = "INSERT INTO rep_dot_cm_det " & _
                         "(bpronro, sec,subsec1,subsec2,subsec3, secdesc,n1, n2, n3, n4, n5, n1desc, n2desc, n3desc, n4desc, n5desc, m1) " & _
                         " VALUES " & _
                         "(" & NroProceso & "," & Seccion & ",-1,-1,-1,'" & NombreSeccion(Seccion) & "'," & _
                         Coordenadas(1) & "," & Coordenadas(2) & "," & Coordenadas(3) & "," & Coordenadas(4) & ", " & _
                         Coordenadas(5) & ",'" & NomNivel(1) & "','" & NomNivel(2) & "','" & NomNivel(3) & "','" & _
                         NomNivel(4) & "','" & NomNivel(5) & "'," & Round(porcentaje) & ")"

                objConn.Execute StrSql, , adExecuteNoRecords
                'Obtiene el ID del registro
                nroCabRep = getLastIdentity(objConn, "rep_dot_cm_det")
                
            Else
                'Arma la cadena para actualizar los valores de los meses restantes
                SqlPorcentajeMes = SqlPorcentajeMes & "m" & NroMes & "= " & Round(porcentaje) & " , "
            End If
            
            porcentaje = 0
        Next  'Sig. NroMes
                
        'Actualiza los meses del 2 al 12 y el campo total ya que solo inserto el m1 y los otros datos
        total = Round(total / MesCorte, 2)
        StrSql = "UPDATE rep_dot_cm_det SET " & SqlPorcentajeMes & "total= " & total & _
                 " WHERE repnro = " & nroCabRep
        objConn.Execute StrSql, , adExecuteNoRecords
                
        Progreso = Progreso + IncPorc
        Call actualizar_progreso(Progreso)
        total = 0
        SqlPorcentajeMes = ""
        
        Call SigConjuntoIndice(Coordenadas, False)
    Next  'Sig. CantElementos
    
                    
        Call CalcularTotalSubTotal(Seccion, TotalSeccion, True, -1, -1, -1)
                    
End Sub

'Calcula la cantidad de empleados segun las estructuras del reporte y el filtro
Sub CalcularPlantillaActiva(Seccion As Integer)
 Dim rsdet As New ADODB.Recordset
 Dim FechaHasta As String
 Dim Campos As String
 Dim valores As String
 Dim j, I As Integer
 Dim EstructuraNivel As String
 Dim Coordenadas() As Integer
 Dim CantEmpleados As Integer
 Dim NroMes As Integer
 Dim nroCabRep As Long
 Dim SqlCantEmpleadoMes As String
 Dim TotalSeccion(13) As Double
 Dim SqlSexo As String
 
    'Inicializa variables
    Call SigConjuntoIndice(Coordenadas, True)
    IncPorc = (22.5 / cantElementos)
    
    'Recorre tantas veces como convinaciones haya (cantElementos)
    For I = 0 To cantElementos - 1
    
        For NroMes = 1 To MesCorte
            
            FechaHasta = ultimo_dia_mes(NroMes, Anio)
            Flog.writeline Espacios(Tabulador * 1) & "Calculando la plantilla activa a la fecha:" & FechaHasta
        
            StrSql = " SELECT count(distinct empleado) cantidad " & _
                     "FROM empleado " & _
                     "INNER JOIN tercero ON empleado.ternro= tercero.ternro " & _
                     "INNER JOIN fases ON empleado.ternro = fases.empleado " & _
                     "INNER JOIN his_estructura as Emp ON Emp.ternro = empleado.ternro AND Emp.tenro = 10 "
                               
            'Filtra la empresa si se selecciono 1 sino no la filtra
            If Empresas <> -1 Then
                StrSql = StrSql & "AND Emp.estrnro = " & Empresas
                StrSql = StrSql & "AND Emp.htetdesde <= " & ConvFecha(FechaHasta) & " AND ((Emp.htethasta >=" & ConvFecha(FechaHasta) & ") OR Emp.htethasta is null)"
            End If
                                
            For j = 1 To UBound(arrTipoEstructura)
                If arrTipoEstructura(j) <> -1 Then
                    StrSql = StrSql & " INNER JOIN his_estructura as he" & j & " ON he" & j & ".ternro = empleado.ternro AND he" & j & ".tenro = " & arrTipoEstructura(j)
                    StrSql = StrSql & " AND he" & j & ".htetdesde <= " & ConvFecha(FechaHasta) & " AND ((he" & j & ".htethasta >=" & ConvFecha(FechaHasta) & ") OR he" & j & ".htethasta is null)"
                    EstructuraNivel = EstructuraNivel & "AND he" & j & ".estrnro IN ( " & arrEstructura(j, Coordenadas(j)) & ") "
                    NomNivel(j) = NombreNiveles(j, Coordenadas(j))
                Else
                    NomNivel(j) = NombreNiveles(j, Coordenadas(j))
                    SqlSexo = "AND tersex= " & arrEstructura(j, Coordenadas(j))
                End If
            Next
                       
            StrSql = StrSql & " WHERE altfec <= " & ConvFecha(FechaHasta)
            'StrSql = StrSql & " AND (bajfec is null or bajfec >=" & ConvFecha(FechaHasta) & ")"
            StrSql = StrSql & " AND (bajfec is null or bajfec >" & ConvFecha(FechaHasta) & ")"
            StrSql = StrSql & EstructuraNivel & " AND empleado in " & empleados
            StrSql = StrSql & SqlSexo
            
            OpenRecordset StrSql, rsdet
            EstructuraNivel = ""
            
            If Not rsdet.EOF Then
                CantEmpleados = rsdet!Cantidad
                If CantEmpleados = 0 Then
                    Flog.writeline Espacios(Tabulador * 2) & "sin empleados, " & StrSql
                End If
            Else
                CantEmpleados = 0
                Flog.writeline Espacios(Tabulador * 2) & "sin empleados, " & StrSql
            End If
        
            TotalSeccion(NroMes) = TotalSeccion(NroMes) + CantEmpleados
            
            Flog.writeline Espacios(Tabulador * 1) & "Empleado de la Plantilla activa: " & CantEmpleados
            'Flog.writeline Espacios(Tabulador * 1) & StrSql
                                    
            If NroMes = 1 Then
                Campos = "(bpronro, sec, subsec1,subsec2,subsec3, secdesc,n1, n2, n3, n4, n5, n1desc, n2desc, n3desc, n4desc, n5desc, m1)"
                valores = "(" & NroProceso & "," & Seccion & ",-1,-1,-1,'" & NombreSeccion(Seccion) & "'," & _
                          Coordenadas(1) & "," & Coordenadas(2) & "," & Coordenadas(3) & "," & Coordenadas(4) & "," & _
                          Coordenadas(5) & ",'" & NomNivel(1) & "','" & NomNivel(2) & "','" & NomNivel(3) & "','" & NomNivel(4) & "','" & NomNivel(5) & "'," & _
                          CantEmpleados & ")"
                                
                StrSql = " INSERT INTO rep_dot_cm_det " & Campos & " VALUES " & valores
                objConn.Execute StrSql, , adExecuteNoRecords
                nroCabRep = getLastIdentity(objConn, "rep_dot_cm_det")
            Else
                'Arma la cadena para actualizar los valores de los meses restantes
                SqlCantEmpleadoMes = SqlCantEmpleadoMes & "m" & NroMes & "= " & CantEmpleados & " , "
            End If
        Next  'Sig NroMes
        
        
        'Actualiza los meses del 2 al 12 y el campo total ya que solo inserto el m1 y los otros datos
        'El Campo total es igual al valor del ultimo mes (12)
        StrSql = "UPDATE rep_dot_cm_det SET " & SqlCantEmpleadoMes & "total= " & CantEmpleados & _
                 " WHERE repnro = " & nroCabRep
        objConn.Execute StrSql, , adExecuteNoRecords
                
        Progreso = Progreso + IncPorc
        actualizar_progreso (Progreso)
        SqlCantEmpleadoMes = ""
        
        Call SigConjuntoIndice(Coordenadas, False)
    Next  'Sig. CantElementos
    
    Call CalcularTotalSubTotal(Seccion, TotalSeccion, True, -1, -1, -1)
    

Exit Sub

MError:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub

'Esta Seccion va en 0 por Ahora
Private Sub CalcularSeccionTres(Seccion As Integer)
 Dim SubSeccion(2) As String
 Dim I, j As Integer
 Dim X As Integer
 Dim TotalSeccion(13) As Double
    
    'Inicializa variables
    Call SigConjuntoIndice(Coordenadas, True)
    SubSeccion(1) = "Colaboradores"   'no va
    SubSeccion(2) = "Becarios"        'no va
    IncPorc = (22.5 / (cantElementos * 2))
    
    'Recorre las Secciones
    For j = 1 To UBound(SubSeccion)
    
        'Recorre tantas veces como convinaciones haya (cantElementos)
        For I = 0 To cantElementos - 1
        
            For X = 1 To UBound(arrTipoEstructura)
                NomNivel(X) = NombreNiveles(X, Coordenadas(X))
            Next
            'Inserta los subtotales
            StrSql = "Insert INTO rep_dot_cm_det (bpronro,sec,subsec1,subsec2,subsec3,secdesc,subsec1desc,n1,n2,n3,n4,n5,n1desc,n2desc,n3desc,n4desc,n5desc," & _
                     "m1,m2,m3,m4,m5,m6,m7,m8,m9,m10,m11,m12,total) VALUES (" & NroProceso & "," & Seccion & "," & j & ",-1,-1,'"
            StrSql = StrSql & NombreSeccion(Seccion) & "','" & SubSeccion(j) & "'," & Coordenadas(1) & ","
            StrSql = StrSql & Coordenadas(2) & "," & Coordenadas(3) & "," & Coordenadas(4) & "," & Coordenadas(5) & ",'" & NomNivel(1) & "','" & NomNivel(2) & "','" & NomNivel(3) & "','" & NomNivel(4) & "','" & NomNivel(5) & "',"
            StrSql = StrSql & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0
            StrSql = StrSql & "," & 0 & "," & 0 & "," & 0 & "," & 0
            StrSql = StrSql & "," & 0 & "," & 0 & "," & 0 & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Progreso = Progreso + IncPorc
            actualizar_progreso (Progreso)
            Call SigConjuntoIndice(Coordenadas, False)
            
        Next  'Sig CantElementos
        
        If j = 1 Then
            Call CalcularTotalSubTotal(Seccion, TotalSeccion, False, j, -1, -1, SubSeccion(j))
        Else
            Call CalcularTotalSubTotal(Seccion, TotalSeccion, True, j, -1, -1, SubSeccion(j))
        End If
        
    
    Next  'Sig SubSeccionN1

Exit Sub

MError:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
End Sub

'EAM- Calcula la Seccion Rotación
Private Sub CalcularRotacion(Seccion As Integer)
 Dim I As Integer
 Dim SubSeccionN1(2) As String
 
 SubSeccionN1(1) = "Altas mes en curso"
 SubSeccionN1(2) = "Bajas mes en curso"
 
    For I = 1 To UBound(SubSeccionN1)
        Select Case I
            Case 1:
                Call CalcularRotSubSeccionAltasMes(Seccion, I, SubSeccionN1(I))
            Case 2:
                Call CalcularRotSubSeccionBajasMes(Seccion, I, SubSeccionN1(I))
        End Select
    Next

End Sub

'EAM- Calcula la Subseccion - Alta del mes
Private Sub CalcularRotSubSeccionAltasMes(Seccion As Integer, SubSeccionN1 As Integer, NomSubSeccionN1 As String)
 Dim rsdet As New ADODB.Recordset
 Dim SubSeccionN2(2) As String
 Dim j As Integer
 Dim I, k As Integer
 Dim FechaDesde, FechaHasta As String
 Dim EstructuraNivel As String
 Dim CantEmpleados As Integer
 Dim SqlCantEmpleadoMes As String
 Dim NroMes As Integer
 Dim TotalRegistro As Double
 Dim TotalSeccion(13) As Double
 Dim Campos, valores As String
 Dim nroCabRep As Long
 Dim SqlSexo As String
 
    'Inicializa variables
    Call SigConjuntoIndice(Coordenadas, True)
    SubSeccionN2(1) = "Procedentes del Grupo"     'no va
    SubSeccionN2(2) = "Procedentes del Exterior"  'no va
    IncPorc = (10 / (cantElementos * 2))
    
    'Recorre las SubSecciones de nivel 1
    For j = 1 To UBound(SubSeccionN2)
        Select Case j
        
        'Procedentes del Grupo
        Case 1:
                        
            'Inicializa variables
            Call SigConjuntoIndice(Coordenadas, True)
    
    
            'Recorre tantas veces como convinaciones haya (cantElementos)
            For I = 0 To cantElementos - 1
    
                For NroMes = 1 To MesCorte
                    FechaDesde = primer_dia_mes(NroMes, Anio)
                    FechaHasta = ultimo_dia_mes(NroMes, Anio)
                    Flog.writeline Espacios(Tabulador * 1) & "Altas menusales ala fecha:" & FechaHasta
        
                    StrSql = "SELECT count(distinct empleado) cantidad " & _
                             "FROM empleado " & _
                             "INNER JOIN tercero ON empleado.ternro= tercero.ternro " & _
                             "INNER JOIN fases ON empleado.ternro = fases.empleado " & _
                             "INNER JOIN his_estructura as Emp ON Emp.ternro = empleado.ternro AND Emp.tenro = 10 "
                               
                    'Filtra la empresa si se selecciono 1 sino no la filtra
                    If Empresas <> -1 Then
                        StrSql = StrSql & " AND Emp.estrnro = " & Empresas
                        StrSql = StrSql & " AND Emp.htetdesde <= " & ConvFecha(FechaHasta) & " AND ((Emp.htethasta >=" & ConvFecha(FechaHasta) & ") OR Emp.htethasta is null)"
                    End If
                                
                    For k = 1 To UBound(arrTipoEstructura)
                        If arrTipoEstructura(k) <> -1 Then
                            StrSql = StrSql & " INNER JOIN his_estructura as he" & k & " ON he" & k & ".ternro = empleado.ternro AND he" & k & ".tenro = " & arrTipoEstructura(k)
                            StrSql = StrSql & " AND he" & k & ".htetdesde <= " & ConvFecha(FechaHasta) & " AND ((he" & k & ".htethasta >=" & ConvFecha(FechaHasta) & ") OR he" & k & ".htethasta is null)"
                            EstructuraNivel = EstructuraNivel & "AND he" & k & ".estrnro IN ( " & arrEstructura(k, Coordenadas(k)) & ") "
                            NomNivel(k) = NombreNiveles(k, Coordenadas(k))
                        Else
                            NomNivel(k) = NombreNiveles(k, Coordenadas(k))
                            SqlSexo = "AND tersex= " & arrEstructura(k, Coordenadas(k))
                        End If
                    Next
                       
                    StrSql = StrSql & " WHERE altfec <= " & ConvFecha(FechaHasta)
                    StrSql = StrSql & " AND altfec >=" & ConvFecha(FechaDesde) & " "
                    StrSql = StrSql & EstructuraNivel & " AND empleado in " & empleados
                    StrSql = StrSql & SqlSexo
                        
                    OpenRecordset StrSql, rsdet
                    EstructuraNivel = ""
                    
                    If Not rsdet.EOF Then
                        CantEmpleados = rsdet!Cantidad
                    Else
                        CantEmpleados = 0
                    End If
        
                    TotalSeccion(NroMes) = TotalSeccion(NroMes) + CantEmpleados
                    TotalRegistro = TotalRegistro + CantEmpleados
            
                    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de altas:: " & CantEmpleados
                    Flog.writeline Espacios(Tabulador * 1) & StrSql
                                    
                    If NroMes = 1 Then
                        Campos = "(bpronro, sec, subsec1,subsec2,subsec3, secdesc,subsec1desc,subsec2desc,n1, n2, n3, n4, n5, n1desc, n2desc, n3desc, n4desc, n5desc, m1)"
                        valores = "(" & NroProceso & "," & Seccion & "," & SubSeccionN1 & "," & j & ",-1,'" & NombreSeccion(Seccion) & "','" & NomSubSeccionN1 & "','" & SubSeccionN2(j) & "'," & _
                                  Coordenadas(1) & "," & Coordenadas(2) & "," & Coordenadas(3) & "," & Coordenadas(4) & "," & _
                                  Coordenadas(5) & ",'" & NomNivel(1) & "','" & NomNivel(2) & "','" & NomNivel(3) & "','" & NomNivel(4) & "','" & NomNivel(5) & "'," & _
                                  CantEmpleados & ")"
                                        
                        StrSql = " INSERT INTO rep_dot_cm_det " & Campos & " VALUES " & valores
                        objConn.Execute StrSql, , adExecuteNoRecords
                        nroCabRep = getLastIdentity(objConn, "rep_dot_cm_det")
                    Else
                        'Arma la cadena para actualizar los valores de los meses restantes
                        SqlCantEmpleadoMes = SqlCantEmpleadoMes & "m" & NroMes & "= " & CantEmpleados & " , "
                    End If
                Next
        
                'Actualiza los meses del 2 al 12 y el campo total ya que solo inserto el m1 y los otros datos
                StrSql = "UPDATE rep_dot_cm_det SET " & SqlCantEmpleadoMes & "total= " & TotalRegistro & _
                         " WHERE repnro = " & nroCabRep
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Progreso = Progreso + IncPorc
                actualizar_progreso (Progreso)
                TotalRegistro = 0
                SqlCantEmpleadoMes = ""
                
                Call SigConjuntoIndice(Coordenadas, False)
            Next
                
            'Seteo el total porque se calcula con los totales de cada registro
            TotalSeccion(13) = 0
            Call CalcularTotalSubTotal(Seccion, TotalSeccion, False, SubSeccionN1, j, -1, NomSubSeccionN1, SubSeccionN2(j))
            
        '________________________________________________________________________________________________________________________
        'Procedentes del Exterior
        Case 2:
            'Recorre tantas veces como convinaciones haya (cantElementos)
            For I = 0 To cantElementos - 1
                For k = 1 To UBound(arrTipoEstructura)
                    NomNivel(k) = NombreNiveles(k, Coordenadas(k))
                Next
                'Inserta los subtotales
                StrSql = "Insert INTO rep_dot_cm_det (bpronro,sec,subsec1,subsec2,subsec3,secdesc,subsec1desc,subsec2desc,n1,n2,n3,n4,n5,n1desc,n2desc,n3desc,n4desc,n5desc," & _
                         "m1,m2,m3,m4,m5,m6,m7,m8,m9,m10,m11,m12,total) VALUES (" & NroProceso & "," & Seccion & "," & SubSeccionN1 & "," & j & ",-1,'"
                StrSql = StrSql & NombreSeccion(Seccion) & "','" & NomSubSeccionN1 & "','" & SubSeccionN2(j) & "'," & Coordenadas(1) & ","
                StrSql = StrSql & Coordenadas(2) & "," & Coordenadas(3) & "," & Coordenadas(4) & "," & Coordenadas(5) & ",'" & NomNivel(1) & "','" & NomNivel(2) & "','" & NomNivel(3) & "','" & NomNivel(4) & "','" & NomNivel(5) & "',"
                StrSql = StrSql & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0
                StrSql = StrSql & "," & 0 & "," & 0 & "," & 0 & "," & 0
                StrSql = StrSql & "," & 0 & "," & 0 & "," & 0 & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                                                
                Progreso = Progreso + IncPorc
                actualizar_progreso (Progreso)
                Call SigConjuntoIndice(Coordenadas, False)
                                
            Next
            
            Call CalcularTotalSubTotal(Seccion, TotalSeccion, True, SubSeccionN1, j, -1, NomSubSeccionN1, SubSeccionN2(j))
       
        End Select
    Next
    
    
    
End Sub


Private Sub CalcularRotSubSeccionBajasMes(Seccion As Integer, SubSeccionN1 As Integer, NomSubSeccionN1 As String)
 Dim rsdet As New ADODB.Recordset
 Dim SubSeccionN2(3) As String
 Dim j As Integer
 Dim I, k As Integer
 Dim FechaDesde, FechaHasta As String
 Dim EstructuraNivel As String
 Dim CantEmpleados As Integer
 Dim SqlCantEmpleadoMes As String
 Dim NroMes As Integer
 Dim TotalSeccion(13) As Double
 Dim TotalRegistro As Double
 Dim Campos, valores As String
 Dim nroCabRep As Long
 Dim SqlSexo As String
 
    'Inicializa variables
    Call SigConjuntoIndice(Coordenadas, True)
    SubSeccionN2(1) = "Voluntarias Destino Grupo"     'no va
    SubSeccionN2(2) = "Voluntarias destino Exterior"  'no va
    SubSeccionN2(3) = "Bajas Forzadas"  'no va
    IncPorc = (10 / (cantElementos * 3))
    
    'Recorre las SubSecciones de nivel 1
    For j = 1 To UBound(SubSeccionN2)
        Select Case j
            
        '______________________________________________________________________________________________________________________
        Case 1: 'Voluntarias Destino Grupo
                        
            'Inicializa variables
            Call SigConjuntoIndice(Coordenadas, True)
        
            'Recorre tantas veces como convinaciones haya (cantElementos)
            For I = 0 To cantElementos - 1
    
                For NroMes = 1 To MesCorte
                    FechaDesde = primer_dia_mes(NroMes, Anio)
                    FechaHasta = ultimo_dia_mes(NroMes, Anio)
                    Flog.writeline Espacios(Tabulador * 1) & "Bajas a la fecha:" & FechaHasta
        
                    StrSql = "SELECT count(distinct empleado) cantidad " & _
                             "FROM empleado " & _
                             "INNER JOIN tercero ON empleado.ternro= tercero.ternro " & _
                             "INNER JOIN fases ON empleado.ternro = fases.empleado " & _
                             "INNER JOIN his_estructura as Emp ON Emp.ternro = empleado.ternro AND Emp.tenro = 10 "
                               
                    'Filtra la empresa si se selecciono 1 sino no la filtra
                    If Empresas <> -1 Then
                        StrSql = StrSql & " AND Emp.estrnro = " & Empresas
                        StrSql = StrSql & " AND Emp.htetdesde <= " & ConvFecha(FechaHasta) & " AND ((Emp.htethasta >=" & ConvFecha(FechaHasta) & ") OR Emp.htethasta is null)"
                    End If
                                
                    For k = 1 To UBound(arrTipoEstructura)
                        If arrTipoEstructura(k) <> -1 Then
                            StrSql = StrSql & " INNER JOIN his_estructura as he" & k & " ON he" & k & ".ternro = empleado.ternro AND he" & k & ".tenro = " & arrTipoEstructura(k)
                            StrSql = StrSql & " AND he" & k & ".htetdesde <= " & ConvFecha(FechaHasta) & " AND ((he" & k & ".htethasta >=" & ConvFecha(FechaHasta) & ") OR he" & k & ".htethasta is null)"
                            EstructuraNivel = EstructuraNivel & "AND he" & k & ".estrnro IN ( " & arrEstructura(k, Coordenadas(k)) & ") "
                            NomNivel(k) = NombreNiveles(k, Coordenadas(k))
                        Else
                            NomNivel(k) = NombreNiveles(k, Coordenadas(k))
                            SqlSexo = "AND tersex= " & arrEstructura(k, Coordenadas(k))
                        End If
                    Next
                       
                    StrSql = StrSql & " WHERE bajfec <= " & ConvFecha(FechaHasta)
                    'StrSql = StrSql & " AND bajfec >=" & ConvFecha(FechaDesde) & " "
                    StrSql = StrSql & " AND bajfec >=" & ConvFecha(FechaDesde) & " AND caunro NOT IN (" & CausaDesp & ") "
                    StrSql = StrSql & EstructuraNivel & " AND empleado in " & empleados
                    StrSql = StrSql & SqlSexo
                    
                    OpenRecordset StrSql, rsdet
                    EstructuraNivel = ""
                    
                    If Not rsdet.EOF Then
                        CantEmpleados = rsdet!Cantidad
                    Else
                        CantEmpleados = 0
                    End If
        
                    TotalSeccion(NroMes) = TotalSeccion(NroMes) + CantEmpleados
                    TotalRegistro = TotalRegistro + CantEmpleados
            
                    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de bajas " & SubSeccionN2(j) & ": " & CantEmpleados
                    Flog.writeline Espacios(Tabulador * 1) & StrSql
                                    
                    If NroMes = 1 Then
                        Campos = "(bpronro, sec, subsec1,subsec2, subsec3,secdesc,subsec1desc,subsec2desc,n1, n2, n3, n4, n5, n1desc, n2desc, n3desc, n4desc, n5desc, m1)"
                        valores = "(" & NroProceso & "," & Seccion & "," & SubSeccionN1 & "," & j & ",-1,'" & NombreSeccion(Seccion) & "','" & NomSubSeccionN1 & "','" & SubSeccionN2(j) & "'," & _
                                  Coordenadas(1) & "," & Coordenadas(2) & "," & Coordenadas(3) & "," & Coordenadas(4) & "," & _
                                  Coordenadas(5) & ",'" & NomNivel(1) & "','" & NomNivel(2) & "','" & NomNivel(3) & "','" & NomNivel(4) & "','" & NomNivel(5) & "'," & _
                                  CantEmpleados & ")"
                                        
                        StrSql = " INSERT INTO rep_dot_cm_det " & Campos & " VALUES " & valores
                        objConn.Execute StrSql, , adExecuteNoRecords
                        nroCabRep = getLastIdentity(objConn, "rep_dot_cm_det")
                    Else
                        'Arma la cadena para actualizar los valores de los meses restantes
                        SqlCantEmpleadoMes = SqlCantEmpleadoMes & "m" & NroMes & "= " & CantEmpleados & " , "
                    End If
                Next
        
                'Actualiza los meses del 2 al 12 y el campo total ya que solo inserto el m1 y los otros datos
                StrSql = "UPDATE rep_dot_cm_det SET " & SqlCantEmpleadoMes & "total= " & TotalRegistro & _
                         " WHERE repnro = " & nroCabRep
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Progreso = Progreso + IncPorc
                actualizar_progreso (Progreso)
                TotalRegistro = 0
                SqlCantEmpleadoMes = ""
        
                Call SigConjuntoIndice(Coordenadas, False)
            Next
            
            'Seteo el total porque se calcula con los totales de cada registro
            TotalSeccion(13) = 0
            Call CalcularTotalSubTotal(Seccion, TotalSeccion, False, SubSeccionN1, j, -1, NomSubSeccionN1, SubSeccionN2(j))
            
        '_______________________________________________________________________________________________________________________
        Case 2: 'Voluntarias Destino Exterior
            'Recorre tantas veces como convinaciones haya (cantElementos)
            For I = 0 To cantElementos - 1
                
                For k = 1 To UBound(arrTipoEstructura)
                    NomNivel(k) = NombreNiveles(k, Coordenadas(k))
                Next
                'Inserta los subtotales
                StrSql = "Insert INTO rep_dot_cm_det (bpronro,sec,subsec1,subsec2,subsec3,secdesc,subsec1desc,subsec2desc,n1,n2,n3,n4,n5,n1desc,n2desc,n3desc,n4desc,n5desc," & _
                         "m1,m2,m3,m4,m5,m6,m7,m8,m9,m10,m11,m12,total) VALUES (" & NroProceso & "," & Seccion & "," & SubSeccionN1 & "," & j & ",-1,'" & _
                         NombreSeccion(Seccion) & "','" & NomSubSeccionN1 & "','" & SubSeccionN2(j) & "'," & Coordenadas(1) & ","
                StrSql = StrSql & Coordenadas(2) & "," & Coordenadas(3) & "," & Coordenadas(4) & "," & Coordenadas(5) & ",'" & NomNivel(1) & "','" & NomNivel(2) & "','" & NomNivel(3) & "','" & NomNivel(4) & "','" & NomNivel(5) & "',"
                StrSql = StrSql & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0
                StrSql = StrSql & "," & 0 & "," & 0 & "," & 0 & "," & 0
                StrSql = StrSql & "," & 0 & "," & 0 & "," & 0 & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Progreso = Progreso + IncPorc
                Call actualizar_progreso(Progreso)
                Call SigConjuntoIndice(Coordenadas, False)
                
            Next
    
             Call CalcularTotalSubTotal(Seccion, TotalSeccion, False, SubSeccionN1, j, -1, NomSubSeccionN1, SubSeccionN2(j))
    
        '_______________________________________________________________________________________________________________________
        Case 3: 'Bajas Forzadas
                        
            'Inicializa variables
            Call SigConjuntoIndice(Coordenadas, True)
            
            'Recorre tantas veces como convinaciones haya (cantElementos)
            For I = 0 To cantElementos - 1
    
                For NroMes = 1 To MesCorte
                    FechaDesde = primer_dia_mes(NroMes, Anio)
                    FechaHasta = ultimo_dia_mes(NroMes, Anio)
                    Flog.writeline Espacios(Tabulador * 1) & "Bajas Forzadas a la fecha:" & FechaHasta
        
                    StrSql = "SELECT count(distinct empleado) cantidad " & _
                             "FROM empleado " & _
                             "INNER JOIN tercero ON empleado.ternro= tercero.ternro " & _
                             "INNER JOIN fases ON empleado.ternro = fases.empleado " & _
                             "INNER JOIN his_estructura as Emp ON Emp.ternro = empleado.ternro AND Emp.tenro = 10 "
                               
                    'Filtra la empresa si se selecciono 1 sino no la filtra
                    If Empresas <> -1 Then
                        StrSql = StrSql & " AND Emp.estrnro = " & Empresas
                        StrSql = StrSql & " AND Emp.htetdesde <= " & ConvFecha(FechaHasta) & " AND ((Emp.htethasta >=" & ConvFecha(FechaHasta) & ") OR Emp.htethasta is null)"
                    End If
                                
                    For k = 1 To UBound(arrTipoEstructura)
                        If arrTipoEstructura(k) <> -1 Then
                            StrSql = StrSql & " INNER JOIN his_estructura as he" & k & " ON he" & k & ".ternro = empleado.ternro AND he" & k & ".tenro = " & arrTipoEstructura(k)
                            StrSql = StrSql & " AND he" & k & ".htetdesde <= " & ConvFecha(FechaHasta) & " AND ((he" & k & ".htethasta >=" & ConvFecha(FechaHasta) & ") OR he" & k & ".htethasta is null)"
                            EstructuraNivel = EstructuraNivel & "AND he" & k & ".estrnro IN ( " & arrEstructura(k, Coordenadas(k)) & ") "
                            NomNivel(k) = NombreNiveles(k, Coordenadas(k))
                        Else
                            NomNivel(k) = NombreNiveles(k, Coordenadas(k))
                            SqlSexo = "AND tersex= " & arrEstructura(k, Coordenadas(k))
                        End If
                    Next
                       
                    StrSql = StrSql & " WHERE bajfec <= " & ConvFecha(FechaHasta)
                    StrSql = StrSql & " AND bajfec >=" & ConvFecha(FechaDesde) & " AND caunro IN (" & CausaDesp & ") "
                    StrSql = StrSql & EstructuraNivel & " AND empleado in " & empleados
                    StrSql = StrSql & SqlSexo
                    
                    OpenRecordset StrSql, rsdet
                    EstructuraNivel = ""
                    
                    If Not rsdet.EOF Then
                        CantEmpleados = rsdet!Cantidad
                    Else
                        CantEmpleados = 0
                    End If
        
                    TotalSeccion(NroMes) = TotalSeccion(NroMes) + CantEmpleados
                    TotalRegistro = TotalRegistro + CantEmpleados
                        
                    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de bajas " & SubSeccionN2(j) & ": " & CantEmpleados
                    Flog.writeline Espacios(Tabulador * 1) & StrSql
                                    
                    If NroMes = 1 Then
                        Campos = "(bpronro, sec, subsec1,subsec2,subsec3,secdesc,subsec1desc,subsec2desc,n1, n2, n3, n4, n5, n1desc, n2desc, n3desc, n4desc, n5desc, m1)"
                        valores = "(" & NroProceso & "," & Seccion & "," & SubSeccionN1 & "," & j & ",-1,'" & NombreSeccion(Seccion) & "','" & NomSubSeccionN1 & "','" & SubSeccionN2(j) & "'," & _
                                  Coordenadas(1) & "," & Coordenadas(2) & "," & Coordenadas(3) & "," & Coordenadas(4) & "," & _
                                  Coordenadas(5) & ",'" & NomNivel(1) & "','" & NomNivel(2) & "','" & NomNivel(3) & "','" & NomNivel(4) & "','" & NomNivel(5) & "'," & _
                                  CantEmpleados & ")"
                                        
                        StrSql = " INSERT INTO rep_dot_cm_det " & Campos & " VALUES " & valores
                        objConn.Execute StrSql, , adExecuteNoRecords
                        nroCabRep = getLastIdentity(objConn, "rep_dot_cm_det")
                    Else
                        'Arma la cadena para actualizar los valores de los meses restantes
                        SqlCantEmpleadoMes = SqlCantEmpleadoMes & "m" & NroMes & "= " & CantEmpleados & " , "
                    End If
                Next
        
                'Actualiza los meses del 2 al 12 y el campo total ya que solo inserto el m1 y los otros datos
                StrSql = "UPDATE rep_dot_cm_det SET " & SqlCantEmpleadoMes & "total= " & TotalRegistro & _
                         " WHERE repnro = " & nroCabRep
                objConn.Execute StrSql, , adExecuteNoRecords
                
                'Setea las variables
                Progreso = Progreso + IncPorc
                Call actualizar_progreso(Progreso)
                TotalRegistro = 0
                SqlCantEmpleadoMes = ""
        
                Call SigConjuntoIndice(Coordenadas, False)
            Next
    
            Call CalcularTotalSubTotal(Seccion, TotalSeccion, True, SubSeccionN1, j, -1, NomSubSeccionN1, SubSeccionN2(j))
    
        End Select
    Next
    
    
    
End Sub

'---------------------------------------------------------------------------------------------------
' procedimiento que busca los empleados que se configuro en el filtro del reporte
'---------------------------------------------------------------------------------------------------

Sub filtro_empleados(ByRef StrSql As String, ByVal Fecha As Date)

Dim StrAgencia As String
Dim StrSelect As String
Dim strjoin As String
Dim StrOrder As String
Dim fecdes As String
Dim fechas As String
Dim rsfiltro As New ADODB.Recordset

On Error GoTo ME_armarsql

StrSql = ""
StrSelect = ""
strjoin = ""
StrOrder = ""

' Busco todos los empleados que cumplen con el filtro

StrAgencia = "" ' cuando queremos todos los empleados

If agencia = "-1" Then
    StrAgencia = " AND empleado.ternro NOT IN (SELECT ternro FROM his_estructura agencia "
    StrAgencia = StrAgencia & " WHERE agencia.tenro=28 AND (agencia.htetdesde<=" & ConvFecha(Fecha)
    StrAgencia = StrAgencia & "     AND (agencia.htethasta IS NULL OR agencia.htethasta>=" & ConvFecha(Fecha) & ")) )"
Else
    If agencia = "-2" Then
        StrAgencia = " AND empleado.ternro IN (SELECT ternro FROM his_estructura agencia "
        StrAgencia = StrAgencia & " WHERE agencia.tenro=28 AND (agencia.htetdesde<=" & ConvFecha(Fecha)
        StrAgencia = StrAgencia & " AND (agencia.htethasta IS NULL OR agencia.htethasta>=" & ConvFecha(Fecha) & ")) )"
    Else
        If agencia <> "0" Then 'este caso se da cuando selecionamos una agencia determinada
            StrAgencia = " AND empleado.ternro in (SELECT ternro from his_estructura agencia "
            StrAgencia = StrAgencia & " WHERE agencia.tenro=28 and agencia.estrnro=" & agencia
            StrAgencia = StrAgencia & "  AND (agencia.htetdesde<=" & ConvFecha(Fecha)
            StrAgencia = StrAgencia & "  AND (agencia.htethasta IS NULL OR agencia.htethasta>=" & ConvFecha(Fecha) & ")) )"
        End If
    End If
End If
 
 
 
If tenro1 <> 0 Then  ' Cuando solo selecionamos el primer nivel
    
    
    strjoin = strjoin & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
    strjoin = strjoin & "  AND (estact1.htetdesde<=" & ConvFecha(Fecha) & " AND (estact1.htethasta IS NULL OR estact1.htethasta>=" & ConvFecha(Fecha) & "))"
    If estrnro1 <> 0 Then
        strjoin = strjoin & " AND estact1.estrnro =" & estrnro1
    End If
    strjoin = strjoin & " INNER JOIN estructura estructura1 ON estructura1.estrnro=estact1.estrnro "
    
    If StrOrder <> "" Then
        StrOrder = StrOrder & ", "
    End If
    StrOrder = StrOrder & " tenro1, estrnro1 "

End If

If tenro2 <> 0 Then  ' ocurre cuando se selecciono hasta el segundo nivel

    
    strjoin = strjoin & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2
    strjoin = strjoin & " AND (estact2.htetdesde<=" & ConvFecha(Fecha) & " AND (estact2.htethasta IS NULL OR estact2.htethasta>=" & ConvFecha(Fecha) & "))"
    If estrnro2 <> 0 Then
        strjoin = strjoin & " AND estact2.estrnro =" & estrnro2
    End If
    strjoin = strjoin & " INNER JOIN estructura estructura2 ON estructura2.estrnro=estact2.estrnro "
    
    If StrOrder <> "" Then
        StrOrder = StrOrder & ", "
    End If
    StrOrder = StrOrder & " tenro2, estrnro2 "

End If

If tenro3 <> 0 Then  ' esto ocurre solo cuando se seleccionan los tres niveles


    strjoin = strjoin & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro =" & tenro3
    strjoin = strjoin & "   AND (estact3.htetdesde<=" & ConvFecha(Fecha) & " AND (estact3.htethasta IS NULL OR  estact3.htethasta>=" & ConvFecha(Fecha) & "))"
    If estrnro3 <> 0 Then 'cuando se le asigna un valor al nivel 3
        strjoin = strjoin & " AND estact3.estrnro =" & estrnro3
    End If
    strjoin = strjoin & " INNER JOIN estructura estructura3 ON estructura3.estrnro=estact3.estrnro "

    If StrOrder <> "" Then
        StrOrder = StrOrder & ", "
    End If
    StrOrder = StrOrder & " tenro3, estrnro3 "
    
End If

                      
StrSql = " SELECT DISTINCT empleado.ternro  "   '  empleado.empest, tercero.tersex,
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " INNER JOIN tercero ON empleado.ternro = tercero.ternro "
StrSql = StrSql & strjoin
StrSql = StrSql & " WHERE " & filtro & StrAgencia

OpenRecordset StrSql, rsfiltro

empleados = "(0"

While Not rsfiltro.EOF
    empleados = empleados & "," & rsfiltro!ternro
    rsfiltro.MoveNext
Wend
 
empleados = empleados & ")"


Exit Sub


ME_armarsql:
    Flog.writeline " Error: Armar consulta del Filtro.- " & Err.Description
    Flog.writeline " Búsqueda de empleados filtrados: " & StrSql
    
    
End Sub

'Obtiene la primera fecha del mes y anio dado
Function primer_dia_mes(mes As Integer, Anio As Integer) As Date
Dim aux As String
     
    primer_dia_mes = C_Date("01/" & mes & "/" & Anio)
    
End Function

'Obtiene la ultima fecha del mes y anio dado
Function ultimo_dia_mes(mes As Integer, Anio As Integer) As Date

Dim mes_sgt As Integer
Dim anio_sgt As Integer

    If mes = 12 Then
        mes_sgt = 1
        anio_sgt = Anio + 1
    Else
        mes_sgt = mes + 1
        anio_sgt = Anio
    End If
    
    ultimo_dia_mes = DateAdd("d", -1, primer_dia_mes(mes_sgt, anio_sgt))
    
End Function

'Actualiza el estado el proceso
Sub actualizar_progreso(Progreso As Double)

TiempoAcumulado = GetTickCount

StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
StrSql = StrSql & " WHERE bpronro = " & NroProceso
objConn.Execute StrSql, , adExecuteNoRecords
    
End Sub

