Attribute VB_Name = "interEKI"
Option Explicit
 
' ---------------------------------------------------------------------------------------------
' Descripcion: Interfaz que genera la tabla EKIfiliacion. La misma se usa para generar una vista que es usada por TimeWare
' Autor      : Juan Pablo Brzozowski
' Fecha      : 06/01/2011

' Modificaciones: 14/02/2011 - Brzozowski Juan Pablo - Se incorporaron los campos tipo_documento y nro_documento en
'                                                      la interface


' ---------------------------------------------------------------------------------------------
 

'--------------------------------------------------------------------------------------
'Datos del proceso
'--------------------------------------------------------------------------------------
Global Const NombreProceso = "EKIinterfaz"
Global Const Ejecutable = "EKIinterfaz.exe"
Global Const Version = "1.00"
Global Const TipoDeProceso = "286"

'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------

Dim fs, f

Dim NroLinea As Long
Dim crpNro As Long
Dim RegLeidos As Long
Dim RegError As Long
Dim RegFecha As Date
Dim NroProceso As Long

Global Path As String
Global NArchivo As String
Global Rta
Global HuboErrores As Boolean
Global EmpErrores As Boolean
Global EmpleadosProcesados As Long
Global CantEmpModificados As Long
Global CantEmpInsertados As Long

 


Private Sub Main()

Dim NombreArchivo As String
Dim Directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim objRs As New ADODB.Recordset
Dim objRs2 As New ADODB.Recordset
Dim objRs3 As New ADODB.Recordset
Dim fechadesde
Dim fechahasta
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim listapronro
Dim Pronro
Dim Ternro
Dim EstadoEmp
Dim arrpronro
Dim rsEmpl As New ADODB.Recordset
 
Dim i
Dim totalEmpleados
Dim cantRegistros
Dim PID As String
Dim tituloReporte As String
Dim parametros As String
Dim ArrParametros
Dim strTempo As String
Dim orden

 

    On Error GoTo CE
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
    

    '------------------------------------------------------------------------
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas 'Modulo mdlDataAccess
    '------------------------------------------------------------------------
 
    
    TiempoInicialProceso = GetTickCount

    Nombre_Arch = PathFLog & "InterfazEKI" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    Flog.writeline " "
    Flog.writeline "Inicio Proceso para la generación de la Interfaz EKI : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
   
    'Obtengo el Process ID
      
    PID = GetCurrentProcessId 'Modulo MdlGlobal
    Flog.writeline "-----------------------------------------------------------------"
    
    Flog.writeline "Nombre del Proceso = " & NombreProceso
    Flog.writeline "Nombre del Ejecutable = " & Ejecutable
    Flog.writeline "Version = " & Version
    Flog.writeline "Tipo De Proceso = " & TipoDeProceso

    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline

    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If

    HuboErrores = False
    
    
    'Obtengo la cantidad de empledos a procesar de la tabla batch_proceso (Son los empleados filtrados previamente)
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    cantRegistros = CLng(objRs!bprcempleados)
    totalEmpleados = cantRegistros
    
    objRs.Close
   
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0, bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
      '---------------------------------------
       'Obtengo los parametros de batch_proceso.bprcparam
      '---------------------------------------
       parametros = objRs!bprcparam
       'Flog.writeline "Parametros del proceso: " & parametros
       
      '---------------------------------------------------------
       
       'EMPIEZA EL PROCESO
  
       'Obtengo los empleados sobre los que tengo que generar los datos de filiacion (Vista 1)
       Flog.writeline "Obtengo los empleados sobre los que tengo que generar los datos de filiacion"
       Call CargarEmpleados(NroProceso, rsEmpl)
       
       Flog.writeline "Inicializo progreso"
       StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
       objConn.Execute StrSql, , adExecuteNoRecords
       
       
       Flog.writeline " "
       Flog.writeline " "
       Flog.writeline "Genero por cada empleado una entrada en la tabla EKIfiliacion"
       Flog.writeline "###################################################################"
       'Genero por cada empleado una entrada en la Vista 1
       
       CantEmpModificados = 0
       CantEmpInsertados = 0
       EmpleadosProcesados = 0
       
       Do Until rsEmpl.EOF
          
          arrpronro = Split(listapronro, ",")
          EmpErrores = False
          Ternro = rsEmpl!Ternro
          EstadoEmp = rsEmpl!Estado
                    
          'GENERO LA VISTA 1
           If Not EmpleadoEnBatch(NroProceso, Ternro) Then
            
            Call InsertarEnBatch(NroProceso, Ternro) 'Inserto el empleado en batch_empleado
            
            Flog.writeline "Tercero " & Ternro
            Flog.writeline "Genero la Vista 1"
            Call generarVista1(NroProceso, Ternro, EstadoEmp)
            Flog.writeline " "
            Flog.writeline "--------------------------------------------------------------------------------------------"
            EmpleadosProcesados = EmpleadosProcesados + 1
           End If
         
          'GENERO LA VISTA 2
          'Esta opcion es para generar la vista 2 por medio de un proceso
          'Flog.writeline "Genero la Vista 2"
          'Call generarVista2(NroProceso, Ternro)
         
        rsEmpl.MoveNext
        '---------------------------------------------------------------------------------------------------------------
        'ACTUALIZO EL PROGRESO DEL PROCESO
        TiempoAcumulado = GetTickCount
        cantRegistros = cantRegistros - 1
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                    ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                    ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
           
        objConn.Execute StrSql, , adExecuteNoRecords
        '---------------------------------------------------------------------------------------------------------------
 
    Loop
    'Elimino las entradas de la tabla batch_empleado con bpronro = NroProceso
    LimpiarBatch (NroProceso)
Else
    Exit Sub
End If
   
'ACTUALIZO EL ESTADO DEL PROCESO
If Not HuboErrores Then
  StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
Else
  StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
End If
    
objConn.Execute StrSql, , adExecuteNoRecords

Flog.writeline "###################################################################"
Flog.writeline "Cantidad de empleados Procesados: " & EmpleadosProcesados
Flog.writeline "Cantidad de empleados Insertados en EKIfiliacion: " & CantEmpInsertados
Flog.writeline "Cantidad de empleados Modificados en EKIfiliacion: " & CantEmpModificados
Flog.writeline "Fin :" & Now
Flog.writeline "###################################################################"


Flog.Close
rsEmpl.Close


Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline "************************************************************"
    Flog.writeline " Error: " & Err.Description & Now
    Flog.writeline " Ultimo sql ejecutado: " & StrSql
    Flog.writeline "************************************************************"
End Sub







'--------------------------------------------------------------------
' Se encarga de generar los datos de la vista 1
'--------------------------------------------------------------------
Sub generarVista1(Pronro, Ternro, EstEmp)
' ---------------------------------------------------------------------------------------------
' Descripcion: Genera una tabla EKIfiliacion con los datos del Empleado dado por el Ternro y el proceso Pronro
' Autor      : JPB
' Fecha      : 06/01/2011
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim cliqnro
Dim profecini As String

'Variables donde se guardan los datos del INSERT final

Dim Legajo                ' Legajo del empleado
Dim Apellido              ' Apellido del empleado
Dim nombre                ' Nombre del empleado
Dim FechaNacimiento       ' Fecha de nacimiento del empleado
Dim Sexo                  ' Codigo del sexo del empleado
Dim Estado                ' Estado del empleado (Activo/Inactivo)
Dim FechaUltimaFase       ' Fecha de la ultima fase del empleado
Dim CodigoConvenio        ' Codigo del convenio laboral del empleado (Codigo de nexus y descripcion)
Dim EmpFecAlta            ' Fecha de ingreso a la empresa
Dim LugarTrabajo           ' Lugar de Pago
Dim PosicionOrganigrama   ' Posicion en el organigrama del empleado
Dim CentroCosto           ' Centro de costos
Dim Puesto                ' Puesto de trabajo del empleado
Dim Funcion               ' Función de la posición ocupada por el empleado
Dim Categoria             ' Categoría laboral o escalafón del empleado
Dim FecAltaCategoria      ' Fecha del último cambio de categoría del empleado
Dim TipoContrato          ' Modalidad de contrato del empleado
Dim RazonSocial           ' Empresa (razón social) a la que pertenece el empleado
Dim TipoDocumento         ' Tipo de documento del empleado
Dim NumeroDocumento       ' Numero de documento del empleado
 

On Error GoTo MError

'-----------------------------------------------------------------
'Inicializo los valores de las variables
'-----------------------------------------------------------------
  Legajo = "null"
  Apellido = "null"
  nombre = "null"
  FechaNacimiento = "null"
  Sexo = "null"
 
   Estado = EstEmp
 
  FechaUltimaFase = "null"
  CodigoConvenio = "null"
  EmpFecAlta = "null"
  LugarTrabajo = "null"
  PosicionOrganigrama = "null"
  CentroCosto = "null"
  Puesto = "null"
  Funcion = "null"
  Categoria = "null"
  FecAltaCategoria = "null"
  TipoContrato = "null"
  RazonSocial = "null"
  TipoDocumento = "null"
  NumeroDocumento = "null"
 
'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
Flog.writeline "Busco los datos del empleado ternro=" & Ternro

StrSql = " SELECT empleado.empleg,empleado.terape,empleado.terape2,empleado.ternom,empleado.ternom2, empleado.empest, "
StrSql = StrSql & " empleado.empfaltagr, tercero.terfecnac, tercero.tersex   "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & "INNER JOIN tercero ON tercero.ternro = empleado.ternro"
StrSql = StrSql & " WHERE empleado.ternro= " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    
   Legajo = rsConsult!EmpLeg
   Apellido = rsConsult!terape & " " & rsConsult!terape2
   nombre = rsConsult!ternom & " " & rsConsult!ternom2
   FechaNacimiento = ConvFecha(rsConsult!terfecnac)
   Sexo = rsConsult!tersex
   'Estado = rsConsult!empest
   EmpFecAlta = ConvFecha(rsConsult!empfaltagr)

Else
   Flog.writeline "Error al obtener los datos del empleado"
   'GoTo MError
End If


'------------------------------------------------------------------
'Busco la fecha de la ultima Fase
'------------------------------------------------------------------
StrSql = " SELECT * FROM fases WHERE empleado = " & Ternro & " ORDER BY altfec DESC"
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   If Not EsNulo(rsConsult!bajfec) Then
     FechaUltimaFase = ConvFecha(rsConsult!bajfec)
   Else
      FechaUltimaFase = ConvFecha(rsConsult!altfec)
   End If
Else
   Flog.writeline "Error al obtener la fecha de la ultima Fase"
   'GoTo MError
End If


'------------------------------------------------------------------
'Busco el Codigo del convenio laboral del empleado
'------------------------------------------------------------------

        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = 19"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           CodigoConvenio = rsConsult!estrnro
        Else
           Flog.writeline "Atencion! No se pudo obtener el Convenio del empleado"
           'GoTo MError
        End If
 

'------------------------------------------------------------------
'Busco el lugar de trabajo del empleado
'------------------------------------------------------------------

 
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = 1"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           LugarTrabajo = rsConsult!estrnro
        Else
           Flog.writeline "Atencion! No se pudo obtener el lugar de pago del empleado"
           'GoTo MError
        End If
 
 
'------------------------------------------------------------------
'Posicion en el organigrama del empleado
'------------------------------------------------------------------

 'Este campo se deja en null



'------------------------------------------------------------------
'Busco el Centro de costos
'------------------------------------------------------------------
  
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = 5"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           CentroCosto = rsConsult!estrnro
        Else
           Flog.writeline "Atencion! No se pudo obtener el centro de costos del empleado"
           'GoTo MError
        End If
 
  
'------------------------------------------------------------------
'Busco Puesto de trabajo del empleado
'------------------------------------------------------------------
  
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = 4"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           Puesto = rsConsult!estrnro
        Else
           Flog.writeline "Atencion! No se pudo obtener el Puesto del empleado"
           'GoTo MError
        End If
 
'------------------------------------------------------------------
'Busco la Función de la posición ocupada por el empleado
'------------------------------------------------------------------
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = 4"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           Funcion = rsConsult!estrnro
        Else
           Flog.writeline "Atencion! No se pudo obtener el Puesto del empleado"
           'GoTo MError
        End If
 
 
'------------------------------------------------------------------
'Busco la Categoria del empleado
'------------------------------------------------------------------
  
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = 3"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           Categoria = rsConsult!estrnro
        Else
           Flog.writeline "Atencion! No se pudo obtener la Categoria del empleado"
           'GoTo MError
        End If
 
 
'------------------------------------------------------------------
'Busco la Fecha del último cambio de categoría del empleado
'------------------------------------------------------------------
  
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = 3 ORDER BY htetdesde DESC"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           
          If Not EsNulo(rsConsult!htetdesde) Then
             FecAltaCategoria = ConvFecha(rsConsult!htetdesde)
          Else
            
            Flog.writeline "Atencion! No se pudo obtener la Fecha del ultimo cambio de categoria del empleado"
            'GoTo MError
          End If
        
        Else
           Flog.writeline "Atencion! No se pudo obtener la Fecha del ultimo cambio de categoria del empleado"
           'GoTo MError
        End If
 
  
  
'------------------------------------------------------------------
'Busco la Modalidad de contrato del empleado
'------------------------------------------------------------------
  
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = 18"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           TipoContrato = rsConsult!estrnro
        Else
           Flog.writeline "Atencion! No se pudo obtener la Modalidad de contrato del empleado"
           'GoTo MError
        End If
  
  
'------------------------------------------------------------------
'Busco la Empresa (razón social) a la que pertenece el empleado
'------------------------------------------------------------------
  
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = 10"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           RazonSocial = rsConsult!estrnro
        Else
           Flog.writeline "Atencion! No se pudo obtener la Empresa (razon social) a la que pertenece el empleado"
           'GoTo MError
        End If
 
'------------------------------------------------------------------
'Busco el Tipo y Numero de Documento del empleado
'------------------------------------------------------------------
  
        StrSql = " SELECT docu.nrodoc AS nrodoc, tipodocu.tidnro AS tipodoc "
        StrSql = StrSql & " FROM tercero "
        StrSql = StrSql & " LEFT JOIN ter_doc docu ON (docu.ternro= tercero.ternro and docu.tidnro>0 and docu.tidnro<5) "
        StrSql = StrSql & " LEFT JOIN tipodocu ON tipodocu.tidnro= docu.tidnro "
        StrSql = StrSql & " WHERE tercero.ternro= " & Ternro
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
        
           If rsConsult!tipodoc <> "" Then
              TipoDocumento = rsConsult!tipodoc
           Else
              Flog.writeline "Atencion! No se pudo obtener el Tipo de Documento del empleado"
           End If
           
           If rsConsult!nrodoc <> "" Then
              NumeroDocumento = rsConsult!nrodoc
           Else
              Flog.writeline "Atencion! No se pudo obtener el Numero de Documento del empleado"
           End If
           
        Else
           Flog.writeline "Atencion! No se pudo obtener el Tipo y Numero de Documento del empleado"
           'GoTo MError
        End If
        
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'Armo la SQL para guardar los datos
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
If (Not EmpleadoEnEKIfiliacion(Legajo)) And (EstEmp = -1) Then
'Si el legajo a incorporar en la tabla no esta cargado y ademas esta activo, hago un insert
    StrSql = " INSERT INTO EKIfiliacion "
    StrSql = StrSql & " (empleg,terape,ternom,terfecnac,tersex,empest,fecultfase,codconvenio,empfecalta,  "
    StrSql = StrSql & " lugarpago,posorg,centrocosto,puesto,funcion,categoria,fecaltcat,tipocont,razsoc,tipo_documento,nro_documento)"
    StrSql = StrSql & " VALUES "
    StrSql = StrSql & "(" & Legajo
    StrSql = StrSql & ",'" & Mid(Apellido, 1, 50) & "'"
    StrSql = StrSql & ",'" & Mid(nombre, 1, 50) & "'"
    StrSql = StrSql & "," & FechaNacimiento
    StrSql = StrSql & "," & Sexo
    StrSql = StrSql & "," & Estado
    StrSql = StrSql & "," & FechaUltimaFase
    StrSql = StrSql & "," & CodigoConvenio
    StrSql = StrSql & "," & EmpFecAlta
    StrSql = StrSql & "," & LugarTrabajo
    StrSql = StrSql & "," & PosicionOrganigrama
    StrSql = StrSql & "," & CentroCosto
    StrSql = StrSql & "," & Puesto
    StrSql = StrSql & "," & Funcion
    StrSql = StrSql & "," & Categoria
    StrSql = StrSql & "," & FecAltaCategoria
    StrSql = StrSql & "," & TipoContrato
    StrSql = StrSql & "," & RazonSocial
    StrSql = StrSql & "," & TipoDocumento
    StrSql = StrSql & ",'" & Mid(NumeroDocumento, 1, 20) & "'"
     
    StrSql = StrSql & ")"
    
    CantEmpInsertados = CantEmpInsertados + 1
End If

If EmpleadoEnEKIfiliacion(Legajo) Then 'Si el legajo a incorporar en la tabla YA esta cargado, hago un UPDATE
    StrSql = " UPDATE EKIfiliacion SET "
    StrSql = StrSql & " terape = '" & Mid(Apellido, 1, 50) & "'"
    StrSql = StrSql & ", ternom = '" & Mid(nombre, 1, 50) & "'"
    StrSql = StrSql & ", terfecnac = " & FechaNacimiento
    StrSql = StrSql & ", tersex = " & Sexo
    StrSql = StrSql & ", empest = " & Estado
    StrSql = StrSql & ", fecultfase = " & FechaUltimaFase
    StrSql = StrSql & ", codconvenio = " & CodigoConvenio
    StrSql = StrSql & ", empfecalta = " & EmpFecAlta
    StrSql = StrSql & ", lugarpago = " & LugarTrabajo
    StrSql = StrSql & ", posorg = " & PosicionOrganigrama
    StrSql = StrSql & ", centrocosto = " & CentroCosto
    StrSql = StrSql & ", puesto = " & Puesto
    StrSql = StrSql & ", funcion = " & Funcion
    StrSql = StrSql & ", categoria = " & Categoria
    StrSql = StrSql & ", fecaltcat = " & FecAltaCategoria
    StrSql = StrSql & ", tipocont = " & TipoContrato
    StrSql = StrSql & ", razsoc = " & RazonSocial
    StrSql = StrSql & ", tipo_documento = " & TipoDocumento
    StrSql = StrSql & ", nro_documento = '" & Mid(NumeroDocumento, 1, 20) & "'"
    StrSql = StrSql & " WHERE empleg = " & Legajo
    
    CantEmpModificados = CantEmpModificados + 1
End If

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

Flog.writeline "SQL : " & StrSql

objConn.Execute StrSql, , adExecuteNoRecords
 
 
Exit Sub

MError:
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error en empleado: " & Legajo & " Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Última SQL ejecutada: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    HuboErrores = True
    EmpErrores = True
End Sub




'JPB - Inserta en la tabla batch_empleado el empleado dado por el numero Tercero y el proceso NroProceso
 Sub InsertarEnBatch(NroProceso, Tercero)
    Dim sql As String
    
    On Error GoTo MError
     
    'INICIO DE TRANSACCION
    objConn.BeginTrans
    
    sql = "INSERT INTO batch_empleado (bpronro, ternro, estado, progreso, beparam) VALUES (" & NroProceso & ", " & Tercero & ", null, null, null)"
    objConn.Execute sql, , adExecuteNoRecords
       
    'FINALIZO LA TRANSACCION
    objConn.CommitTrans
 

Exit Sub

MError:
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error en empleado: " & Legajo & " Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Última SQL ejecutada: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    HuboErrores = True
    EmpErrores = True

 End Sub



'JPB - Borra de la tabla batch_empleado todos los empleados del proceso NroProceso
Sub LimpiarBatch(NroProceso)
    Dim sql As String
    On Error GoTo MError
     
    'INICIO DE TRANSACCION
    objConn.BeginTrans
    
    sql = "DELETE FROM batch_empleado WHERE bpronro = " & NroProceso
    objConn.Execute sql, , adExecuteNoRecords
       
    'FINALIZO LA TRANSACCION
    objConn.CommitTrans
      
Exit Sub

MError:
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error en empleado: " & Legajo & " Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Última SQL ejecutada: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    HuboErrores = True
    EmpErrores = True

 End Sub
 

 
 

Sub CargarEmpleados_En_batch()
' Descripcion: Busca los empleados ACTIVOS e INCATIVOS (dados de baja en la ultima semana a partir de la fecha actual)
'              y los incorpora en la tabla batch_empleado, e incorpora un nuevo proceso en batch_proceso
' Autor      : JPB
' Fecha      : 06/01/2011
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim Legajo
Dim FechaInicioLicencia
Dim FechaFinLicencia
Dim CodigoLicencia

Dim l_desde_sql
Dim l_hasta_sql
Dim l_id
Dim l_hora
Dim l_dia
Dim l_bpronro

Dim StrSql As String

Dim rsEmp As New ADODB.Recordset
Dim rsBatch As New ADODB.Recordset
 
On Error GoTo MError
      
       'Busco los datos de las novedades horarias del empleado
        StrSql = " SELECT * FROM empleado"
        OpenRecordset StrSql, rsEmp
        
        If rsEmp.EOF Then
         Flog.writeline "La tabla empleado esta vacia "
        Else
         Flog.writeline "Busca los empleados ACTIVOS e INCATIVOS (dados de baja en la ultima semana a partir de la fecha actual) y los incorpora en la tabla batch_empleado, e incorpora un nuevo proceso en batch_proceso"
        End If
        
                  
        'INICIO DE TRANSACCION
        objConn.BeginTrans

        'INSERTO EN BATCH PROCESO
        l_desde_sql = "NULL"
        l_hasta_sql = "NULL"
        l_id = ""  ' Session("Username")
           
        StrSql = " INSERT INTO batch_proceso "
        StrSql = StrSql & " (btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta, bprcparam, "
        StrSql = StrSql & " bprcestado, bprcprogreso, bprcfecfin, bprchorafin, bprctiempo, empnro, bprcempleados,bprcurgente) "
        StrSql = StrSql & " VALUES (91 ," & ConvFecha(Date) & ", '" & l_id & "','" & Format(Now, "hh:mm:ss ") & "' "
        StrSql = StrSql & " , " & l_desde_sql & ", " & l_hasta_sql
        StrSql = StrSql & " , '', 'Pendiente', null , null, null, 0, 0, null,0)"
        
        objConn.Execute StrSql, , adExecuteNoRecords
           
        'OBTENGO EL NRO DE PROCESO
        StrSql = "select * from batch_proceso order by bpronro desc"
        OpenRecordset StrSql, rsBatch
        l_bpronro = rsBatch!Ternro
        rsBatch.Close
        
        'INSERTO LOS EMPLEADOS EN BATCH EMPLEADO
'        Do While Not rsEmp.EOF
'
'           StrSql = "INSERT INTO batch_empleado "
'           StrSql = StrSql & " (bpronro, ternro, estado) "
'           StrSql = StrSql & " VALUES (" & l_bpronro & "," & rsEmp!Ternro & ",null)"
'           objConn.Execute StrSql, , adExecuteNoRecords
'
'           rsEmp.MoveNext
'
'        Loop
        
       rsEmp.Close
       
       'FINALIZO LA TRANSACCION
       objConn.CommitTrans
  
Exit Sub

MError:
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error al iniciar el proceso: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Última SQL ejecutada: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    HuboErrores = True
    EmpErrores = True


End Sub


'JPB - Verifica si el empleado esta en EKIfiliacion dado su legajo
Public Function EmpleadoEnEKIfiliacion(EmpLeg)
  
  Dim rsTabla As New ADODB.Recordset
  Dim sql As String
   
  On Error GoTo MError

  sql = " SELECT * from EKIfiliacion where empleg = " & EmpLeg
  OpenRecordset sql, rsTabla

  If Not rsTabla.EOF Then
    EmpleadoEnEKIfiliacion = True
  Else
    EmpleadoEnEKIfiliacion = False
  End If

rsTabla.Close

Exit Function

MError:
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error : " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Última SQL ejecutada: " & sql
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    HuboErrores = True
    EmpErrores = True
  
End Function



'JPB - Verifica si el empleado+nroproc esta en una EKIfiliacion por su legajo
Public Function EmpleadoEnBatch(NroProceso, Tercero)
  
  Dim rsBatch As New ADODB.Recordset
  Dim sql As String
   
  On Error GoTo MError

  sql = " SELECT * from batch_empleado where bpronro = " & NroProceso & " AND ternro = " & Tercero
  OpenRecordset sql, rsBatch

  If Not rsBatch.EOF Then
    EmpleadoEnBatch = True
  Else
    EmpleadoEnBatch = False
  End If

rsBatch.Close


Exit Function

MError:
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error : " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Última SQL ejecutada: " & sql
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    HuboErrores = True
    EmpErrores = True
  
End Function









Public Sub bus_Anti0(Ternro, pliqanio, pliqmes, ByRef dia As Integer, ByRef mes As Integer, ByRef Anio As Integer)

Dim q As Integer

Dim FechaAux As Date

    Bien = False
    Valor = 0
        
    If pliqmes = 12 Then
        FechaAux = CDate("1/1/" & pliqanio + 1) - 1
    Else
        FechaAux = CDate("01/" & pliqmes + 1 & "/" & pliqanio) - 1
    End If
        
    Call bus_Antiguedad(Ternro, "REAL", FechaAux, dia, mes, Anio, q)
    
End Sub


Public Sub bus_Antiguedad(ByVal Ternro As Integer, ByVal TipoAnt As String, ByVal fechafin As String, ByRef dia As Integer, ByRef mes As Integer, ByRef Anio As Integer, ByRef DiasHabiles As Integer)

Dim aux1 As Long
Dim aux2 As Long
Dim aux3 As Long
Dim fecalta As Date
Dim fecbaja As Date
Dim Seguir As Date
Dim q As Integer

Dim NombreCampo As String

Dim rs_Fases As New ADODB.Recordset

NombreCampo = ""
DiasHabiles = 0

Select Case UCase(TipoAnt)
Case "SUELDO":
    NombreCampo = "sueldo"
Case "INDEMNIZACION":
    NombreCampo = "indemnizacion"
Case "VACACIONES":
    NombreCampo = "vacaciones"
Case "REAL":
    NombreCampo = "real"
Case Else
End Select


' FGZ -27/01/2004
StrSql = "SELECT * FROM fases WHERE empleado = " & Ternro & _
         " AND " & NombreCampo & " = -1 " & _
         " AND not altfec is null " & _
         " AND not (bajfec is null AND estado = 0)" & _
         " AND altfec <= " & ConvFecha(fechafin)

OpenRecordset StrSql, rs_Fases

Do While Not rs_Fases.EOF
    fecalta = rs_Fases!altfec
    
    ' Verificar si se trata de un registro completo (alta/baja) o solo de un alta
    If rs_Fases!Estado Then
        fecbaja = fechafin ' solo es un alta, tomar el fecha-fin
    ElseIf rs_Fases!bajfec <= CDate(fechafin) Then
        fecbaja = rs_Fases!bajfec  ' se trata de un registro completo
    Else
        fecbaja = fechafin ' hasta la fecha ingresada
    End If
    
    Call DIF_FECHAS2(fecalta, fecbaja, aux1, aux2, aux3)
    
    If rs_Fases.RecordCount = 1 Then
        dia = aux1
        mes = aux2
        Anio = aux3
    Else
        dia = dia + aux1
        mes = mes + aux2 + Int(dia / 30)
        Anio = Anio + aux3 + Int(mes / 12)
        dia = dia Mod 30
        mes = mes Mod 12
    End If
        
    If Anio = 0 Then
        Call DiasTrab(Ternro, fecalta, fecbaja, aux1)
        DiasHabiles = DiasHabiles + aux1
    End If
    
siguiente:
    rs_Fases.MoveNext
Loop

If Anio <> 0 Then
    DiasHabiles = 0
End If

' Cierro todo y Libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing

End Sub





Public Sub DiasTrab(ByVal Ternro As Integer, ByVal Desde As Date, ByVal Hasta As Date, ByRef DiasH As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la cantidad de dias trabajados de acuerdo al turno en que se trabaja y
'              de acuerdo a los dias que figuran como feriados en la tabla de feriados.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim d1 As Integer
Dim d2 As Integer
Dim aux As Integer
Dim aux2 As Integer
Dim dxsem As Integer

Dim rs_pais As New ADODB.Recordset
Dim rs_feriados As New ADODB.Recordset

    dxsem = 5
    
    d1 = Weekday(Desde)
    d2 = Weekday(Hasta)
    
    aux = DateDiff("d", Desde, Hasta) + 1
    If aux < 7 Then
        DiasH = Minimo(aux, dxsem)
    Else
        If aux = 7 Then
            DiasH = dxsem
        Else
            aux2 = 8 - d1 + d2
            If aux2 < 7 Then
                aux2 = Minimo(aux2, dxsem)
            Else
                If aux2 = 7 Then
                    aux2 = dxsem
                End If
            End If
            
            If aux2 >= 7 Then
                aux2 = Abs(aux2 - 7) + Int(aux2 / 7) * dxsem
            Else
                aux2 = aux2 + Int((aux2 - aux2) / 7) * dxsem
            End If
        End If
    End If
    
    aux = 0
    
    StrSql = "SELECT * FROM pais INNER JOIN tercero ON tercero.paisnro = pais.paisnro WHERE tercero.ternro = " & Ternro
    OpenRecordset StrSql, rs_pais
    
    If Not rs_pais.EOF Then
        ' Resto los Feriados Nacionales
        StrSql = "SELECT * FROM feriado WHERE tipferinro = 2 " & _
                 " AND fericodext = " & rs_pais!paisnro & _
                 " AND ferifecha >= " & ConvFecha(Desde) & _
                 " AND ferifecha < " & ConvFecha(Hasta)
        OpenRecordset StrSql, rs_feriados
        
        Do While Not rs_feriados.EOF
            If Weekday(rs_feriados!ferifecha) > 1 Then
                DiasH = DiasH - 1
            End If
            
            ' Siguiente Feriado
            rs_feriados.MoveNext
        Loop
    End If


    ' Resto los feriados por Convenio
    StrSql = "SELECT * FROM empleado INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro " & _
             " INNER JOIN fer_estr ON fer_estr.tenro = his_estructura.tenro " & _
             " INNER JOIN feriado ON fer_estr.ferinro = feriado.ferinro " & _
             " WHERE empleado.ternro = " & Ternro & _
             " AND feriado.tipferinro = 2" & _
             " AND feriado.ferifecha >= " & ConvFecha(Desde) & _
             " AND feriado.ferifecha < " & ConvFecha(Hasta)
    OpenRecordset StrSql, rs_feriados
    
    Do While Not rs_feriados.EOF
        If Weekday(rs_feriados!ferifecha) > 1 Then
            DiasH = DiasH - 1
        End If
        
        ' Siguiente Feriado
        rs_feriados.MoveNext
    Loop
    
    
    ' cierro todo y libero
    If rs_pais.State = adStateOpen Then rs_pais.Close
    If rs_feriados.State = adStateOpen Then rs_feriados.Close
        
    Set rs_feriados = Nothing
    Set rs_pais = Nothing

End Sub


'--------------------------------------------------------------------
' Se encarga de buscar los empleados a cargar en la tabla EKIfiliacion
'--------------------------------------------------------------------
Sub CargarEmpleados(NroProc, ByRef rsEmpl As ADODB.Recordset)
 
 Dim CantEmpTablaEKI
 Dim StrEmpl As String
 Dim rsEKI As New ADODB.Recordset
 
  StrEmpl = "SELECT * from EKIfiliacion "
  OpenRecordset StrEmpl, rsEKI
  CantEmpTablaEKI = rsEKI.RecordCount
  rsEKI.Close
     
  If CantEmpTablaEKI = 0 Then 'Si la tabla se carga por primera vez, busca los empleados activos
    StrEmpl = "SELECT fases.estado,* from empleado   "
    StrEmpl = StrEmpl & " INNER JOIN fases ON fases.empleado=empleado.ternro  AND (estado=-1) "
  Else
     StrEmpl = "SELECT fases.estado, * from empleado "
     StrEmpl = StrEmpl & " INNER JOIN fases ON fases.empleado=empleado.ternro  ORDER BY empleado.empleg,fases.altfec DESC "
  End If
   
  
  OpenRecordset StrEmpl, rsEmpl
  'EmpleadosProcesados = rsEmpl.RecordCount
 
  
    
End Sub

Function numberForSQL(Str)
   
  numberForSQL = Replace(Str, ",", ".")

End Function

