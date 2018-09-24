Attribute VB_Name = "MdlExportacion"
Option Explicit

Global IdUser As String
Global Fecha As Date
Global Hora As String

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo As Date
Global Fecha_Desde As Date
Global Fecha_Hasta As Date
Global Progreso As Double
Global StrSql2 As String

Global HuboErrores As Boolean
Global EmpErrores As Boolean



Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Proceso.
' Autor      : JMH
' Fecha      : 07/03/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim ArrParametros

    strCmdLine = Command()
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
    OpenConnection strconexion, objconnProgreso
    
    Nombre_Arch = PathFLog & "Exp_Indura" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 79 AND bpronro =" & NroProcesoBatch
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
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' ,bprcprogreso=100 WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.Close
    objconnProgreso.Close
    objConn.Close
    
End Sub
Public Sub GeneracionEmpresa(ByVal Empresa As Integer, ByVal bpronro As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de la Exportacion de Indura
' Autor      : JMH
' Fecha      : 18/03/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim Aux_Total_Importe
Dim strLinea As String
Dim Aux_Linea As String

Dim rs_Periodo As New ADODB.Recordset
Dim rs_Proceso As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset

Dim Cabecera As String

Const ForReading = 1
Const TristateFalse = 0
Dim fExport
Dim Directorio As String
Dim Archivo As String
Dim Intentos As Integer
Dim carpeta
Dim Separador

Dim Aux_str As String
Dim TipoCodEmpresa As String

Dim Cantidad As Long
Dim cantidadProcesada As Long
Dim Progreso As Long

Dim Direccion1 As String
Dim Direccion2 As String
Dim Cuit As String
Dim IngresosBrutos As String

'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 252"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then

    Separador = rs_Modelo!modseparador
    If Not IsNull(rs_Modelo!modarchdefault) Then
        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
    Else
        Flog.writeline "El modelo no tiene configurada la carpeta destino. El archivo será generado en el directorio default"
    End If
Else
    Flog.writeline "No se encontró el modelo. El archivo será generado en el directorio default"
End If

'Activo el manejador de errores
On Error Resume Next

'Archivo para la cabecera del Pedido de Pago
Archivo = Directorio & "\Empresa_Indura.txt"
Set fs = CreateObject("Scripting.FileSystemObject")
Set fExport = fs.CreateTextFile(Archivo, True)

If Err.Number <> 0 Then
    Flog.writeline "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(Directorio)
    Set fExport = fs.CreateTextFile(Archivo, True)
End If
'desactivo el manejador de errores

On Error GoTo 0


' Comienzo la transaccion
MyBeginTrans

StrSql = "SELECT estructura.estrdabr, calle, nro, codigopostal, telnro, locdesc, empresa.empnro, empactiv, "
StrSql = StrSql & " posdes, nrodoc, estrdabr "
StrSql = StrSql & " FROM  estructura "
StrSql = StrSql & " INNER JOIN empresa  ON empresa.estrnro = estructura.estrnro "
StrSql = StrSql & " INNER JOIN tercero  ON tercero.ternro = empresa.ternro "
StrSql = StrSql & " INNER JOIN ter_doc  ON ter_doc.ternro = tercero.ternro AND ter_doc.tidnro = 6 "
StrSql = StrSql & " LEFT JOIN posicion  ON posicion.posnro = tercero.posnro "
StrSql = StrSql & " LEFT JOIN cabdom ON empresa.ternro = cabdom.ternro "
StrSql = StrSql & " LEFT JOIN detdom ON detdom.domnro = cabdom.domnro "
StrSql = StrSql & " LEFT JOIN localidad ON localidad.locnro = detdom.locnro "
StrSql = StrSql & " LEFT JOIN telefono ON telefono.domnro = cabdom.domnro "
StrSql = StrSql & " WHERE estructura.estrnro = " & Empresa
OpenRecordset StrSql, rs_Periodo

Cantidad = rs_Periodo.RecordCount
cantidadProcesada = Cantidad

Dim Error As Boolean

If Not rs_Periodo.EOF Then
           
   Direccion1 = rs_Periodo!calle & " " & rs_Periodo!nro
   Direccion2 = rs_Periodo!codigopostal & " - " & rs_Periodo!locdesc & " - " & rs_Periodo!telnro
   
   Cuit = rs_Periodo!nrodoc
   
   StrSql = " SELECT * "
   StrSql = StrSql & " FROM confrep "
   StrSql = StrSql & " INNER JOIN ter_doc  ON ter_doc.tidnro = confrep.confval "
   StrSql = StrSql & " WHERE confrep.repnro = 123 AND confrep.conftipo = 'TD' "
   OpenRecordset StrSql, rs_Confrep
    
   If Not rs_Confrep.EOF Then
      IngresosBrutos = rs_Periodo!nrodoc
   End If
           
   Cabecera = rs_Periodo!Empnro & Separador & rs_Periodo!estrdabr & Separador & Direccion1 & Separador & Direccion2 & Separador
   Cabecera = Cabecera & rs_Periodo!empactiv & Separador & Cuit & Separador & IngresosBrutos
           
   fExport.writeline Cabecera
        
   TiempoAcumulado = GetTickCount
          
   cantidadProcesada = cantidadProcesada - 1
          
   Progreso = 25
   
   StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
            ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
            ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProcesoBatch
                  
    objConn.Execute StrSql, , adExecuteNoRecords
        
End If

rs_Periodo.Close
fExport.Close

MyCommitTrans

Set rs_Periodo = Nothing
Set rs_Proceso = Nothing

End Sub
Public Sub GeneracionEmpleado(ByVal Periodo As Integer, ByVal Proceso As String, ByVal bpronro As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de la Exportacion de Roche
' Autor      : JMH
' Fecha      : 07/03/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim Aux_Total_Importe
Dim strLinea As String
Dim Aux_Linea As String

Dim rs_Periodo As New ADODB.Recordset
Dim rs_Cabdom As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset

Dim Cabecera As String

Const ForReading = 1
Const TristateFalse = 0
Dim fExport
Dim Directorio As String
Dim Archivo As String
Dim Intentos As Integer
Dim carpeta
Dim Separador As String

Dim Aux_str As String
Dim TipoCodEmpresa As String

Dim Cantidad As Integer
Dim cantidadProcesada As Integer
Dim Progreso As Long

Dim Legajo  As Integer
Dim ApeyNom As String
Dim Direccion As String
Dim Localidad As String
Dim Telefono As String
Dim Documento As String
Dim EstadoCivil As String
Dim Convenio As String
Dim Categoria As String
Dim Cargo As String
Dim AFJP As String
Dim CajaAhorro As String
Dim Cuil As String
Dim sexo As String
Dim Nacionalidad As String
Dim FechaNacimiento As Date
Dim fechaIngreso As Date
Dim FechaAntiguedad As String
Dim AnioAntiguedad  As Integer
Dim Centro As String
Dim Departamento As String
Dim Seccion As String
Dim OSocial As String
Dim Dia As Integer
Dim mes As Integer
Dim anio As Integer


On Error GoTo MError

'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 252"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then

    Separador = rs_Modelo!modseparador
    If Not IsNull(rs_Modelo!modarchdefault) Then
        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
    Else
        Flog.writeline "El modelo no tiene configurada la carpeta destino. El archivo será generado en el directorio default"
    End If
Else
    Flog.writeline "No se encontró el modelo. El archivo será generado en el directorio default"
End If

'Archivo para la cabecera del Pedido de Pago
Archivo = Directorio & "\Empleado_Indura.txt"
Set fs = CreateObject("Scripting.FileSystemObject")
Set fExport = fs.CreateTextFile(Archivo, True)

If Err.Number <> 0 Then
    Flog.writeline "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(Directorio)
    Set fExport = fs.CreateTextFile(Archivo, True)
End If

' Comienzo la transaccion
MyBeginTrans

StrSql = " SELECT empleado.ternro, empleg, periodo.pliqnro, periodo.pliqhasta, periodo.pliqdesde, profecpago, cabliq.cliqnro, "
StrSql = StrSql & " empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
StrSql = StrSql & " FROM  periodo "
StrSql = StrSql & " INNER JOIN proceso  ON proceso.pliqnro = periodo.pliqnro AND proceso.pronro IN (" & Proceso & ")"
StrSql = StrSql & " INNER JOIN cabliq   ON cabliq.pronro = proceso.pronro "
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
StrSql = StrSql & " WHERE periodo.pliqnro = " & Periodo
StrSql = StrSql & " GROUP BY empleado.ternro, empleg, periodo.pliqnro, periodo.pliqhasta, periodo.pliqdesde, profecpago, cabliq.cliqnro, "
StrSql = StrSql & " empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
OpenRecordset StrSql, rs_Periodo

Cantidad = rs_Periodo.RecordCount
cantidadProcesada = Cantidad

Dim Error As Boolean

Do While Not rs_Periodo.EOF
           
   Error = False
   
   'Para el apellido y nombre
   ApeyNom = rs_Periodo!terape & " " & rs_Periodo!terape2 & ", " & rs_Periodo!ternom & " " & rs_Periodo!ternom2
   
   'Para la direccion, localidad, telefono, documento y estado civil
   StrSql = " SELECT calle, nro, piso, oficdepto, locdesc, telnro, tidsigla, nrodoc, estcivdesabr, codigopostal "
   StrSql = StrSql & " FROM  cabdom "
   StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro "
   StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = cabdom.ternro "
   StrSql = StrSql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
   StrSql = StrSql & " INNER JOIN telefono ON detdom.domnro = telefono.domnro AND telefono.teldefault = -1 "
   StrSql = StrSql & " INNER JOIN ter_doc ON ter_doc.ternro = " & rs_Periodo!ternro & " AND ter_doc.tidnro <= 4 "
   StrSql = StrSql & " INNER JOIN tipodocu ON tipodocu.tidnro = ter_doc.tidnro "
   StrSql = StrSql & " INNER JOIN estcivil ON estcivil.estcivnro = tercero.estcivnro "
   StrSql = StrSql & " WHERE cabdom.ternro = " & rs_Periodo!ternro
   OpenRecordset StrSql, rs_Cabdom
    
   If Not rs_Cabdom.EOF Then
      Direccion = rs_Cabdom!calle & " " & rs_Cabdom!nro & " " & rs_Cabdom!piso & " " & rs_Cabdom!oficdepto
      Localidad = rs_Cabdom!codigopostal & " - " & rs_Cabdom!locdesc
      Telefono = rs_Cabdom!telnro
      Documento = rs_Cabdom!tidsigla & " " & rs_Cabdom!nrodoc
      EstadoCivil = rs_Cabdom!estcivdesabr
   Else
      Direccion = ""
      Localidad = ""
      Telefono = ""
      Documento = ""
      EstadoCivil = ""
      Flog.writeline "No se encontraron algunos datos del empleados como la Dirección o Localidad o Teléfono o Documento el convenio del empleado: " & Legajo
   End If
   rs_Cabdom.Close
   
   Legajo = rs_Periodo!empleg
   
   'Para el convenio
   StrSql = " SELECT estrdabr "
   StrSql = StrSql & " FROM  his_estructura "
   StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
   StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Periodo!ternro
   StrSql = StrSql & " AND his_estructura.tenro  = 19 AND (his_estructura.htetdesde <= " & ConvFecha(rs_Periodo!pliqhasta) & ")"
   StrSql = StrSql & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(rs_Periodo!pliqdesde) & ")"
   OpenRecordset StrSql, rs_Estructura
    
   If Not rs_Estructura.EOF Then
      Convenio = rs_Estructura!estrdabr
   Else
      Convenio = ""
      Flog.writeline "No se encontro el Convenio del empleado: " & Legajo
   End If
   rs_Estructura.Close
   
   'Para la categoria
   StrSql = " SELECT estrdabr "
   StrSql = StrSql & " FROM  his_estructura "
   StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
   StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Periodo!ternro
   StrSql = StrSql & " AND his_estructura.tenro  = 3 AND (his_estructura.htetdesde <= " & ConvFecha(rs_Periodo!pliqhasta) & ")"
   StrSql = StrSql & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(rs_Periodo!pliqdesde) & ")"
   OpenRecordset StrSql, rs_Estructura
    
   If Not rs_Estructura.EOF Then
      Categoria = rs_Estructura!estrdabr
   Else
      Categoria = ""
      Flog.writeline "No se encontro la Categoría del empleado: " & Legajo
   End If
   rs_Estructura.Close
   
   'Para el cargo
   StrSql = " SELECT estrdabr "
   StrSql = StrSql & " FROM  his_estructura "
   StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
   StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Periodo!ternro
   StrSql = StrSql & " AND his_estructura.tenro  = 4 AND (his_estructura.htetdesde <= " & ConvFecha(rs_Periodo!pliqhasta) & ")"
   StrSql = StrSql & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(rs_Periodo!pliqdesde) & ")"
   OpenRecordset StrSql, rs_Estructura
    
   If Not rs_Estructura.EOF Then
      Cargo = rs_Estructura!estrdabr
   Else
      Cargo = ""
      Flog.writeline "No se encontró el Cargo del empleado: " & Legajo
   End If
   rs_Estructura.Close
   
   'Para el AFJP
   StrSql = " SELECT estrdabr "
   StrSql = StrSql & " FROM  his_estructura "
   StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
   StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Periodo!ternro
   StrSql = StrSql & " AND his_estructura.tenro  = 15 AND (his_estructura.htetdesde <= " & ConvFecha(rs_Periodo!pliqhasta) & ")"
   StrSql = StrSql & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(rs_Periodo!pliqdesde) & ")"
   OpenRecordset StrSql, rs_Estructura
    
   If Not rs_Estructura.EOF Then
      AFJP = rs_Estructura!estrdabr
   Else
      AFJP = ""
      Flog.writeline "No se encontro la AFJP del empleado: " & Legajo
   End If
   rs_Estructura.Close
   
   'Para la caja de ahorro
   StrSql = " SELECT ctabnro "
   StrSql = StrSql & " FROM ctabancaria "
   StrSql = StrSql & " WHERE ctabancaria.ternro = " & rs_Periodo!ternro & " AND ctabancaria.fpagnro = 2 "
   OpenRecordset StrSql, rs_Estructura
    
   If Not rs_Estructura.EOF Then
      CajaAhorro = rs_Estructura!ctabnro
   Else
      CajaAhorro = ""
      Flog.writeline "No se encontro la Caja de Ahorro del empleado: " & Legajo
   End If
   rs_Estructura.Close
   
   'Para el CUIL
   StrSql = " SELECT nrodoc "
   StrSql = StrSql & " FROM  ter_doc "
   StrSql = StrSql & " WHERE ter_doc.ternro = " & rs_Periodo!ternro & " AND ter_doc.tidnro = 10 "
   OpenRecordset StrSql, rs_Estructura
    
   If Not rs_Estructura.EOF Then
      Cuil = rs_Estructura!nrodoc
   Else
      Cuil = ""
      Flog.writeline "No se encontro el CUIL del empleado: " & Legajo
   End If
   rs_Estructura.Close
   
   'Para el sexo, nacionalidad y para la fecha de nacimiento
   StrSql = " SELECT tersex, nacionalnro, terfecnac "
   StrSql = StrSql & " FROM  tercero "
   StrSql = StrSql & " WHERE tercero.ternro = " & rs_Periodo!ternro
   OpenRecordset StrSql, rs_Estructura
    
   If Not rs_Estructura.EOF Then
      If rs_Estructura!tersex = -1 Then
         sexo = "Masculino"
      Else
         sexo = "Femenino"
      End If
      If rs_Estructura!nacionalnro = 1 Then
         Nacionalidad = "Nativo/a"
      Else
         Nacionalidad = "Extranjero/a"
      End If
   End If
   FechaNacimiento = rs_Estructura!terfecnac
   
   rs_Estructura.Close
   
   'Para la fecha de ingreso
   StrSql = " SELECT empfaltagr "
   StrSql = StrSql & " FROM  empleado "
   StrSql = StrSql & " WHERE empleado.ternro = " & rs_Periodo!ternro
   OpenRecordset StrSql, rs_Estructura
    
   If Not rs_Estructura.EOF Then
      fechaIngreso = rs_Estructura!empfaltagr
   Else
      fechaIngreso = ""
      Flog.writeline "No se encontro la Fecha de Ingreso del empleado: " & Legajo
   End If
   rs_Estructura.Close
   
   
   'Para la fecha de antiguedad
   FechaAntiguedad = ""
   
   'Para los años de antiguedad reconocida
   Call antiguedad(rs_Periodo!ternro, Date, Dia, mes, anio)
   AnioAntiguedad = anio
         
   'Para el Centro
   StrSql = " SELECT estrdabr "
   StrSql = StrSql & " FROM  his_estructura "
   StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
   StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Periodo!ternro
   StrSql = StrSql & " AND his_estructura.tenro  = 5 AND (his_estructura.htetdesde <= " & ConvFecha(rs_Periodo!pliqhasta) & ")"
   StrSql = StrSql & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(rs_Periodo!pliqdesde) & ")"
   OpenRecordset StrSql, rs_Estructura
    
   If Not rs_Estructura.EOF Then
      Centro = rs_Estructura!estrdabr
   Else
      Centro = ""
      Flog.writeline "No se encontro el Centro del empleado: " & Legajo
   End If
   rs_Estructura.Close
   
   'Para la Seccion
   StrSql = " SELECT estrdabr "
   StrSql = StrSql & " FROM  his_estructura "
   StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
   StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Periodo!ternro
   StrSql = StrSql & " AND his_estructura.tenro  = 2 AND (his_estructura.htetdesde <= " & ConvFecha(rs_Periodo!pliqhasta) & ")"
   StrSql = StrSql & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(rs_Periodo!pliqdesde) & ")"
   OpenRecordset StrSql, rs_Estructura
    
   If Not rs_Estructura.EOF Then
       Seccion = rs_Estructura!estrdabr
   Else
       Seccion = ""
       Flog.writeline "No se encontro la Sección del empleado: " & Legajo
   End If
   rs_Estructura.Close
   
   'Para el Departamento
   StrSql = " SELECT estrdabr "
   StrSql = StrSql & " FROM  his_estructura "
   StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
   StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Periodo!ternro
   StrSql = StrSql & " AND his_estructura.tenro  = 9 AND (his_estructura.htetdesde <= " & ConvFecha(rs_Periodo!pliqhasta) & ")"
   StrSql = StrSql & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(rs_Periodo!pliqdesde) & ")"
   OpenRecordset StrSql, rs_Estructura
    
   If Not rs_Estructura.EOF Then
      Departamento = rs_Estructura!estrdabr
   Else
      Departamento = ""
      Flog.writeline "No se encontro el Departamento del empleado: " & Legajo
   End If
   rs_Estructura.Close
   
   'Para la Obra Social
   StrSql = " SELECT estrdabr "
   StrSql = StrSql & " FROM  his_estructura "
   StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
   StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Periodo!ternro
   StrSql = StrSql & " AND his_estructura.tenro  = 17 AND (his_estructura.htetdesde <= " & ConvFecha(rs_Periodo!pliqhasta) & ")"
   StrSql = StrSql & " AND (his_estructura.htethasta is null or his_estructura.htethasta >=" & ConvFecha(rs_Periodo!pliqdesde) & ")"
   OpenRecordset StrSql, rs_Estructura
    
   If Not rs_Estructura.EOF Then
      OSocial = rs_Estructura!estrdabr
   Else
      OSocial = ""
      Flog.writeline "No se encontro la Obra Social del empleado: " & Legajo
   End If
   rs_Estructura.Close
      
   Cabecera = ApeyNom & Separador & Direccion & Separador & Localidad & Separador & Telefono & Separador
   Cabecera = Cabecera & Documento & Separador & EstadoCivil & Separador & Legajo & Separador
   Cabecera = Cabecera & Convenio & Separador & Categoria & Separador & Cargo & Separador & AFJP & Separador & CajaAhorro & Separador
   Cabecera = Cabecera & Cuil & Separador & sexo & Separador & Nacionalidad & Separador & FechaNacimiento & Separador
   Cabecera = Cabecera & fechaIngreso & Separador & FechaAntiguedad & Separador & AnioAntiguedad & Separador & Centro & Separador
   Cabecera = Cabecera & Departamento & Separador & Seccion & Separador & OSocial
        
   fExport.writeline Cabecera
        
   TiempoAcumulado = GetTickCount
          
   cantidadProcesada = cantidadProcesada - 1
          
   Progreso = 25 + Fix((((Cantidad - cantidadProcesada) * 100) / Cantidad) / 4)
   
   StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
            ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
            ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProcesoBatch
                  
   objConn.Execute StrSql, , adExecuteNoRecords
       
   rs_Periodo.MoveNext
Loop

rs_Periodo.Close
fExport.Close

MyCommitTrans

Set rs_Periodo = Nothing
Set rs_Cabdom = Nothing
Set rs_Estructura = Nothing

Exit Sub

MError:
    Flog.writeline "Error en empleado: " & Legajo & " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True

End Sub
Public Sub GeneracionLiquidacion(ByVal Periodo As Integer, ByVal Proceso As String, ByVal bpronro As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de la Exportacion de Indura
' Autor      : JMH
' Fecha      : 21/03/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim Aux_Total_Importe
Dim strLinea As String
Dim Aux_Linea As String

Dim rs_Periodo As New ADODB.Recordset
Dim rs_Proceso As New ADODB.Recordset
Dim rs_PedidoPago As New ADODB.Recordset
Dim rs_AcuLiq As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset

Dim Cabecera As String

Const ForReading = 1
Const TristateFalse = 0
Dim fExport
Dim Directorio As String
Dim Archivo As String
Dim Intentos As Integer
Dim carpeta
Dim Separador As String

Dim Aux_str As String
Dim TipoCodEmpresa As String

Dim Cantidad As Integer
Dim cantidadProcesada As Integer
Dim Progreso As Long

Dim idLiq As String
Dim NroPeriodo As String
Dim Remuneracion As Double
Dim RemuneracionStr As String
Dim Descuento As Double
Dim DescuentoStr As String
Dim ImporteNeto As Double
Dim PeriodoLiq As String
Dim FechaDeposito As Date
Dim Banco As String

Dim Entera As String
Dim Decimales As String
Dim Formateado As String
Dim Reemplazo1 As String
Dim Reemplazo2 As String
Dim Longitud As Integer
Dim RemuneracionForm As String
Dim DescuentoForm As String
Dim ImporteNetoForm As String

'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 252"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    Separador = rs_Modelo!modseparador

    If Not IsNull(rs_Modelo!modarchdefault) Then
        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
    Else
        Flog.writeline "El modelo no tiene configurada la carpeta destino. El archivo será generado en el directorio default"
    End If
Else
    Flog.writeline "No se encontró el modelo. El archivo será generado en el directorio default"
End If

'Activo el manejador de errores
On Error Resume Next

'Archivo para la cabecera del Pedido de Pago
Archivo = Directorio & "\Liquidacion_Indura.txt"
Set fs = CreateObject("Scripting.FileSystemObject")
Set fExport = fs.CreateTextFile(Archivo, True)

If Err.Number <> 0 Then
    Flog.writeline "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(Directorio)
    Set fExport = fs.CreateTextFile(Archivo, True)
End If
'desactivo el manejador de errores

On Error GoTo 0


' Comienzo la transaccion
MyBeginTrans

StrSql = " SELECT DISTINCT empleg, empleado.ternro, periodo.pliqdesc, periodo.pliqhasta, periodo.pliqmes, periodo.pliqanio, periodo.pliqnro "
StrSql = StrSql & " FROM  periodo "
StrSql = StrSql & " INNER JOIN proceso  ON proceso.pliqnro = periodo.pliqnro AND proceso.pronro IN (" & Proceso & ")"
StrSql = StrSql & " INNER JOIN cabliq   ON cabliq.pronro = proceso.pronro "
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
StrSql = StrSql & " WHERE periodo.pliqnro = " & Periodo
StrSql = StrSql & " GROUP BY empleg, empleado.ternro, periodo.pliqdesc, periodo.pliqhasta, periodo.pliqmes, periodo.pliqanio, periodo.pliqnro "
OpenRecordset StrSql, rs_Periodo

Cantidad = rs_Periodo.RecordCount
cantidadProcesada = Cantidad

Dim Error As Boolean

Do While Not rs_Periodo.EOF
           
   Error = False
   
   Remuneracion = 0
   Descuento = 0
   ImporteNeto = 0
              
   'Para la fecha de deposito
   StrSql = " SELECT min(profecpago) as FechaDeposito "
   StrSql = StrSql & " FROM proceso "
   StrSql = StrSql & " WHERE proceso.pronro IN (" & Proceso & ")"
   OpenRecordset StrSql, rs_Proceso
   If Not rs_Proceso.EOF Then
      FechaDeposito = rs_Proceso!FechaDeposito
   End If
   
   'Para el Banco
   StrSql = " SELECT bandesc "
   StrSql = StrSql & " FROM pedidopago "
   StrSql = StrSql & " INNER JOIN pago ON pago.ppagnro = pedidopago.ppagnro AND pago.ternro = " & rs_Periodo!ternro
   StrSql = StrSql & " INNER JOIN banco ON banco.ternro = pedidopago.bannro "
   StrSql = StrSql & " WHERE pedidopago.pliqnro = " & Periodo
   OpenRecordset StrSql, rs_PedidoPago
   
   If Not rs_PedidoPago.EOF Then
      Banco = rs_PedidoPago!bandesc
   End If
   
   StrSql = " SELECT acu_liq.acunro, acu_liq.almonto, acu_liq.alcant "
   StrSql = StrSql & " FROM cabliq "
   StrSql = StrSql & " INNER JOIN  acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
   StrSql = StrSql & " WHERE cabliq.empleado = " & rs_Periodo!ternro & " AND cabliq.pronro IN (" & Proceso & ")"
   OpenRecordset StrSql, rs_AcuLiq
    
   Do While Not rs_AcuLiq.EOF
   
       StrSql = " SELECT * "
       StrSql = StrSql & " FROM confrep "
       StrSql = StrSql & " WHERE confrep.repnro = 123 AND confrep.conftipo = 'AC' AND confrep.confval = " & rs_AcuLiq!acunro
       OpenRecordset StrSql, rs_Confrep
    
       
       If Not rs_Confrep.EOF Then
          If rs_Confrep!confnrocol = 1 Then
             Remuneracion = Remuneracion + rs_AcuLiq!almonto
          ElseIf rs_Confrep!confnrocol = 2 Then
              Descuento = Descuento + rs_AcuLiq!almonto
          ElseIf rs_Confrep!confnrocol = 3 Then
              ImporteNeto = ImporteNeto + rs_AcuLiq!almonto
          End If
       End If
       rs_Confrep.Close
       
       rs_AcuLiq.MoveNext
   Loop
   rs_AcuLiq.Close
   
   PeriodoLiq = CStr(rs_Periodo!pliqmes) & "/" & CStr(rs_Periodo!pliqanio)
                
   Formateado = FormatNumber(Remuneracion, 2)
   Reemplazo1 = Replace(Formateado, ",", "")
   Reemplazo2 = Replace(Reemplazo1, ".", "")
   Longitud = Len(CStr(Reemplazo2)) - 2
   Entera = Mid(Reemplazo2, 1, Longitud)
   Decimales = Right(Reemplazo2, 2)
   Formateado = Entera & "." & Decimales
   RemuneracionForm = Formateado
   
   Formateado = FormatNumber(Descuento, 2)
   Reemplazo1 = Replace(Formateado, ",", "")
   Reemplazo2 = Replace(Reemplazo1, ".", "")
   Longitud = Len(CStr(Reemplazo2)) - 2
   Entera = Mid(Reemplazo2, 1, Longitud)
   Decimales = Right(Reemplazo2, 2)
   Formateado = Entera & "." & Decimales
   DescuentoForm = Formateado
   
   Formateado = FormatNumber(ImporteNeto, 2)
   Reemplazo1 = Replace(Formateado, ",", "")
   Reemplazo2 = Replace(Reemplazo1, ".", "")
   Longitud = Len(CStr(Reemplazo2)) - 2
   Entera = Mid(Reemplazo2, 1, Longitud)
   Decimales = Right(Reemplazo2, 2)
   Formateado = Entera & "." & Decimales
   ImporteNetoForm = Formateado
                
   Cabecera = Format_StrNro(rs_Periodo!empleg, 7, False, "") & Separador & rs_Periodo!pliqdesc & Separador
   Cabecera = Cabecera & rs_Periodo!pliqhasta & Separador & FechaDeposito & Separador & Banco & Separador & PeriodoLiq & Separador
   Cabecera = Cabecera & RemuneracionForm & Separador & DescuentoForm & Separador & ImporteNetoForm & Separador
        
   fExport.writeline Cabecera
   
   TiempoAcumulado = GetTickCount
          
   cantidadProcesada = cantidadProcesada - 1
          
   Progreso = 50 + Fix((((Cantidad - cantidadProcesada) * 100) / Cantidad) / 4)
   
   StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
            ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
            ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProcesoBatch
                  
   objConn.Execute StrSql, , adExecuteNoRecords
    
    
   rs_Periodo.MoveNext
Loop

rs_Periodo.Close
fExport.Close

MyCommitTrans

Set rs_Periodo = Nothing
Set rs_Confrep = Nothing

End Sub
Public Sub GeneracionMovimiento(ByVal Periodo As Integer, ByVal Proceso As String, ByVal bpronro As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de la Exportacion de Roche
' Autor      : JMH
' Fecha      : 07/03/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim Aux_Total_Importe
Dim strLinea As String
Dim Aux_Linea As String

Dim rs_Periodo As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset

Dim Cabecera As String

Const ForReading = 1
Const TristateFalse = 0
Dim fExport
Dim Directorio As String
Dim Archivo As String
Dim Intentos As Integer
Dim carpeta
Dim Separador As String

Dim Aux_str As String
Dim TipoCodEmpresa As String

Dim Cantidad As Integer
Dim cantidadProcesada As Integer
Dim Progreso As Long

Dim idLiq As String
Dim NroPeriodo As String
Dim PercepcionStr As String
Dim PeriodoLiq As String

Dim Entera As String
Dim Decimales As String
Dim Formateado As String
Dim Reemplazo1 As String
Dim Reemplazo2 As String
Dim Longitud As Integer
Dim RemuneracionForm As String
Dim DescuentoForm As String
Dim ImporteNetoForm As String

'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 252"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then

    Separador = rs_Modelo!modseparador
    If Not IsNull(rs_Modelo!modarchdefault) Then
        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
    Else
        Flog.writeline "El modelo no tiene configurada la carpeta destino. El archivo será generado en el directorio default"
    End If
Else
    Flog.writeline "No se encontró el modelo. El archivo será generado en el directorio default"
End If

'Activo el manejador de errores
On Error Resume Next

'Archivo para la cabecera del Pedido de Pago
Archivo = Directorio & "\Movimientos_Indura.txt"
Set fs = CreateObject("Scripting.FileSystemObject")
Set fExport = fs.CreateTextFile(Archivo, True)

If Err.Number <> 0 Then
    Flog.writeline "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(Directorio)
    Set fExport = fs.CreateTextFile(Archivo, True)
End If
'desactivo el manejador de errores

On Error GoTo 0


' Comienzo la transaccion
MyBeginTrans

StrSql = " SELECT empleg, periodo.pliqnro, cabliq.cliqnro, periodo.pliqmes, periodo.pliqanio "
StrSql = StrSql & " FROM  periodo "
StrSql = StrSql & " INNER JOIN proceso  ON proceso.pliqnro = periodo.pliqnro AND proceso.pronro IN (" & Proceso & ")"
StrSql = StrSql & " INNER JOIN cabliq   ON cabliq.pronro = proceso.pronro "
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
StrSql = StrSql & " WHERE periodo.pliqnro = " & Periodo
StrSql = StrSql & " GROUP BY empleg, periodo.pliqnro, cabliq.cliqnro, periodo.pliqmes, periodo.pliqanio "
OpenRecordset StrSql, rs_Periodo

Cantidad = rs_Periodo.RecordCount
cantidadProcesada = Cantidad

Dim Error As Boolean

Do While Not rs_Periodo.EOF
           
   Error = False
   
   StrSql = " SELECT detliq.concnro, detliq.dlimonto, detliq.dlicant, concepto.tconnro, concepto.concabr, CAST (concepto.conccod as int) AS orden "
   StrSql = StrSql & " FROM  detliq "
   StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
   StrSql = StrSql & " WHERE detliq.cliqnro = " & rs_Periodo!cliqnro
   StrSql = StrSql & " ORDER BY orden "
   OpenRecordset StrSql, rs_Detliq
    
   Do While Not rs_Detliq.EOF
   
       PeriodoLiq = CStr(rs_Periodo!pliqmes) & "/" & CStr(rs_Periodo!pliqanio)
       
       Formateado = FormatNumber(rs_Detliq!dlicant, 2)
       Reemplazo1 = Replace(Formateado, ",", "")
       Reemplazo2 = Replace(Reemplazo1, ".", "")
       Longitud = Len(CStr(Reemplazo2)) - 2
       Entera = Mid(Reemplazo2, 1, Longitud)
       Decimales = Right(Reemplazo2, 2)
       Formateado = Entera & "." & Decimales
       DescuentoForm = Formateado
        
       Formateado = FormatNumber(rs_Detliq!dlimonto, 2)
       Reemplazo1 = Replace(Formateado, ",", "")
       Reemplazo2 = Replace(Reemplazo1, ".", "")
       Longitud = Len(CStr(Reemplazo2)) - 2
       Entera = Mid(Reemplazo2, 1, Longitud)
       Decimales = Right(Reemplazo2, 2)
       Formateado = Entera & "." & Decimales
       ImporteNetoForm = Formateado
       
       If rs_Detliq!dlimonto >= 0 Then
          Cabecera = rs_Periodo!empleg & Separador & PeriodoLiq & Separador & rs_Detliq!concabr & Separador
          Cabecera = Cabecera & "" & Separador & DescuentoForm & Separador & "" & Separador
          Cabecera = Cabecera & ImporteNetoForm
          
          Else: Cabecera = rs_Periodo!empleg & Separador & PeriodoLiq & Separador & rs_Detliq!concabr & Separador
          Cabecera = Cabecera & "" & Separador & DescuentoForm & Separador & "" & Separador
          Cabecera = Cabecera & "" & Separador & ImporteNetoForm
          
       End If
       fExport.writeline Cabecera
              
       rs_Detliq.MoveNext
   Loop
   rs_Detliq.Close
   
   TiempoAcumulado = GetTickCount
          
   cantidadProcesada = cantidadProcesada - 1
          
   Progreso = 75 + Fix((((Cantidad - cantidadProcesada) * 100) / Cantidad) / 4)
   If Progreso <= 99 Then
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                 ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProcesoBatch
                  
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
   rs_Periodo.MoveNext
Loop

rs_Periodo.Close
fExport.Close

MyCommitTrans

Set rs_Periodo = Nothing
Set rs_Detliq = Nothing
Set rs_Confrep = Nothing

End Sub
Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : JMH
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim ArrParametros
Dim Exportacion As Integer
Dim Empresa As Integer
Dim Periodo As Integer
Dim Proceso As String

'Orden de los parametros
'Pedido de Ticket

ArrParametros = Split(Parametros, "@")
' Levanto cada parametro por separado
'Exportacion = arrParametros(0)
Periodo = ArrParametros(0)
Proceso = ArrParametros(1)
Empresa = ArrParametros(2)

Call GeneracionEmpresa(Empresa, bpronro)
Call GeneracionEmpleado(Periodo, Proceso, bpronro)
Call GeneracionLiquidacion(Periodo, Proceso, bpronro)
Call GeneracionMovimiento(Periodo, Proceso, bpronro)

End Sub


Public Sub antiguedad(ByVal ternro As Integer, ByVal FechaFin As Date, ByRef Dia As Integer, ByRef mes As Integer, ByRef anio As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la antiguedad al dia de hoy de un empleado en :
'              dias hábiles(si es menor que un año) o en dias, meses y años en caso contrario.
'              antiguedad.p
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim aux1 As Integer
Dim aux2 As Integer
Dim aux3 As Integer
Dim fecalta As Date
Dim fecbaja As Date
Dim seguir As Date
Dim q As Integer

Dim NombreCampo As String

Dim rs_Fases As New ADODB.Recordset

NombreCampo = "real"

' FGZ -27/01/2004
StrSql = "SELECT * FROM fases WHERE empleado = " & ternro & _
         " AND " & NombreCampo & " = -1 " & _
         " AND not altfec is null " & _
         " AND not (bajfec is null AND estado = 0)" & _
         " AND altfec <= " & ConvFecha(FechaFin)

OpenRecordset StrSql, rs_Fases

Dia = 0
mes = 0
anio = 0

Do While Not rs_Fases.EOF
    fecalta = rs_Fases!altfec
        
    ' Verificar si se trata de un registro completo (alta/baja) o solo de un alta
    If rs_Fases!Estado Then
        fecbaja = FechaFin ' solo es un alta, tomar el fecha-fin
    ElseIf rs_Fases!bajfec <= FechaFin Then
        fecbaja = rs_Fases!bajfec  ' se trata de un registro completo
    Else
        fecbaja = FechaFin ' hasta la fecha ingresada
    End If
    
    Call DIF_FECHAS2(fecalta, fecbaja, aux1, aux2, aux3)
    'Call Dif_Fechas(fecalta, fecbaja, aux1, aux2, aux3)
    If rs_Fases.RecordCount = 1 Then
        Dia = aux1
        mes = aux2
        anio = aux3
    Else
        Dia = Dia + aux1
        mes = mes + aux2 + Int(Dia / 30)
        anio = anio + aux3 + Int(mes / 12)
        Dia = Dia Mod 30
        mes = mes Mod 12
    End If
        
siguiente:
    rs_Fases.MoveNext
Loop

' Cierro todo y Libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing

End Sub


