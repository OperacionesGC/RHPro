Attribute VB_Name = "repRecibosPasantes"
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaModificacion = "24/09/2007"
'Global Const UltimaModificacion = " " '24/09/2007 - G. Bauer - se coloco el nro de columna para que vaya buscar el acumulador que corresponda.
'----------------------------------------------------------------------------------
'Global Const Version = "1.02"
'Global Const FechaModificacion = "30/06/2008"
'Global Const UltimaModificacion = " " '30/06/2008 - Lisandro Moro - Se agrego el detalle de la liquidacion a los pasantes, opcional si se configura.
'----------------------------------------------------------------------------------
'Global Const Version = "1.03"
'Global Const FechaModificacion = "23/12/2008"
'Global Const UltimaModificacion = " " 'Martin Ferraro - A partir del recibo de pasante estandar se genero una custom para deloitte I.
'----------------------------------------------------------------------------------
'Global Const Version = "1.04"
'Global Const FechaModificacion = "26/02/2009"
'Global Const UltimaModificacion = " " 'Lisandro Moro  - Se corrigio la busqueda de conceptos.
'                                                       Si no se encuentra el concepto u acumulador, no se guarda en detalle.
'----------------------------------------------------------------------------------
Global Const Version = "1.05"
Global Const FechaModificacion = "31/07/2009"
Global Const UltimaModificacion = " " 'Martin Ferraro  - Encriptacion de string connection
'----------------------------------------------------------------------------------

Dim fs, f
'Global Flog

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
Global arrTipoConc(1000) As Integer
Dim ArrConcAcu(20, 4) ' Lo dimensiono en 20 pero por ahora se pide para utilizar con 5
Dim cantConAcu As Integer
'0=concnrocol, 1=confestiq, 2=conftipo, 3=confval, 4=confval2


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
Dim fechadesde
Dim fechahasta
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim listapronro
Dim pronro
Dim ternro
Dim arrpronro
Dim rsEmpl As New ADODB.Recordset
Dim acunroSueldo
Dim I
Dim totalEmpleados
Dim cantRegistros
Dim PID As String
Dim tituloReporte As String
Dim parametros As String
Dim ArrParametros
Dim strTempo As String
Dim modeloReporte As Long
Dim orden As Long

'    strCmdLine = Command()
'    ArrParametros = Split(strCmdLine, " ", -1)
'    If UBound(ArrParametros) > 0 Then
'        If IsNumeric(ArrParametros(0)) Then
'            NroProceso = ArrParametros(0)
'            Etiqueta = ArrParametros(1)
'        Else
'            Exit Sub
'        End If
'    Else
'        If IsNumeric(strCmdLine) Then
'            NroProceso = strCmdLine
'        Else
'            Exit Sub
'        End If
'    End If

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

    Nombre_Arch = PathFLog & "RecibosPasantes" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    
    tituloReporte = ""

    TiempoInicialProceso = GetTickCount

    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    HuboErrores = False
    
    Flog.writeline "Inicio Proceso de Recibos de Pasantes : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       cantRegistros = CInt(objRs!bprcempleados)
       totalEmpleados = cantRegistros
    
       'Obtengo los parametros del proceso
       parametros = objRs!bprcparam
       ArrParametros = Split(parametros, "@")
       
       'Obtengo la lista de procesos
       listapronro = ArrParametros(0)
       
       'Armo el titulo del reporte
        tituloReporte = ArrParametros(1)
        
       'Armo el modelo del reporte
        modeloReporte = 1
       
       'EMPIEZA EL PROCESO
       ' 24/09/2007 - G. Bauer Se coloca fijo el nro de columna que va a buscar a la configuracion del reporte.
       'Busco en el confrep el numero de cuenta que se va a usar para
       ' buscar el valor del sueldo
       StrSql = " SELECT * FROM confrep "
       StrSql = StrSql & " WHERE repnro = 139 " 'AND confnrocol = " & modeloReporte
       StrSql = StrSql & " ORDER BY confnrocol "
      
       OpenRecordset StrSql, objRs2

        If Not objRs2.EOF Then
            Dim a As Integer
            For a = 0 To UBound(ArrConcAcu, 1)
            'Inicializo la matrix
            '0=concnrocol, 1=confestiq, 2=conftipo, 3=confval, 4=confval2
                ArrConcAcu(a, 0) = 0
                ArrConcAcu(a, 1) = ""
                ArrConcAcu(a, 2) = ""
                ArrConcAcu(a, 3) = 0
                ArrConcAcu(a, 4) = 0
            Next a
            
            a = 0
            cantConAcu = 0
            Do While Not objRs2.EOF
                cantConAcu = cantConAcu + 1
                'En la promet columna SIEMPRE va el NETO
                If objRs2("confnrocol") = 1 Then
                    acunroSueldo = objRs2!confval
                End If
                'objRs2 ("")
                ArrConcAcu(a, 0) = objRs2("confnrocol")
                ArrConcAcu(a, 1) = objRs2("confetiq")
                ArrConcAcu(a, 2) = objRs2("conftipo")
                ArrConcAcu(a, 3) = objRs2("confval")
                ArrConcAcu(a, 4) = objRs2("confval2")
                
                objRs2.MoveNext
                a = a + 1
            Loop
            Flog.writeline "Se cargo el ConfRep"
        Else
            Flog.writeline "No esta configurado el ConfRep"
            Exit Sub
        End If
        
       objRs2.Close
              
       Flog.writeline "Obtengo los datos del confrep"
       
       'Obtengo los empleados sobre los que tengo que generar los recibos
       CargarEmpleados NroProceso, rsEmpl
       
       StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   " WHERE bpronro = " & NroProceso
       
       'Genero por cada empleado un recibo de sueldo
       orden = 0
       Do Until rsEmpl.EOF
          arrpronro = Split(listapronro, ",")
          EmpErrores = False
          ternro = rsEmpl!ternro
          'Genero un recibo de sueldo para el empleado por cada proceso
          For I = 0 To UBound(arrpronro)
             pronro = arrpronro(I)
             Flog.writeline
             Flog.writeline "Generando recibo de pasante empleado " & ternro & " para el proceso " & pronro
             Call generarDatosRecibo(pronro, ternro, acunroSueldo, tituloReporte, orden)
             orden = orden + 1
             'Actualizo el estado del proceso
             TiempoAcumulado = GetTickCount
          Next
          
          cantRegistros = cantRegistros - 1
          
          StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                      ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                      ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
                
          objConn.Execute StrSql, , adExecuteNoRecords
          
          'Si se generaron todos los recibos de sueldo del empleado correctamente lo borro
          If Not EmpErrores Then
              StrSql = " DELETE FROM batch_empleado "
              StrSql = StrSql & " WHERE bpronro = " & NroProceso
              StrSql = StrSql & " AND ternro = " & ternro
    
              objConn.Execute StrSql, , adExecuteNoRecords
          End If
          
          rsEmpl.MoveNext
       Loop
    
    Else
        Exit Sub
    End If
   
    If Not HuboErrores Then
        StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
        StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    
    'Actualizo el estado del proceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now

End Sub

'--------------------------------------------------------------------
' Se encarga de generar los datos
'--------------------------------------------------------------------
Sub generarDatosRecibo(pronro, ternro, acunroSueldo, tituloReporte, orden)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim rsDet As New ADODB.Recordset
Dim cliqnro
Dim profecini As String

'Variables donde se guardan los datos del INSERT final
Dim apellido
Dim nombre
Dim apellido2
Dim nombre2
Dim Legajo
Dim pliqnro
Dim pliqmes
Dim pliqanio
Dim pliqdesc
Dim neto
Dim profecpago
Dim EmpNombre As String
Dim sueldo

'--- detalle ---
Dim repnro As Long
Dim bpronro As Long
'Dim ternro As Long
'Dim pronro As Long
Dim concabr As String
Dim Conccod As String
Dim concnro As Long
Dim dlimonto As Double

Dim Sucursal As String
Dim EmpCuit As String
Dim EmpDire As String
Dim EmpEstrnro As Long
Dim EmpTernro As Long
Dim EmpLogo As String
Dim EmpLogoAlto As Long
Dim EmpLogoAncho As Long
Dim EmpFirma As String
Dim EmpFirmaAlto As Long
Dim EmpFirmaAncho As Long
Dim Cuil As String


On Error GoTo MError
'------------------------------------------------------------------
'Obtengo el nro de cabezera de liquidacion y la fecha de pago del proceso
'------------------------------------------------------------------
StrSql = " SELECT cabliq.cliqnro, proceso.profecpago, proceso.prodesc, proceso.profecini FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso ON cabliq.pronro = proceso.pronro AND proceso.pronro = " & pronro
StrSql = StrSql & " WHERE cabliq.empleado=" & ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   cliqnro = rsConsult!cliqnro
   profecpago = rsConsult!profecpago
Else
   Flog.writeline "Error al obtener los datos del periodo"
   GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empfecalta,empremu "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   nombre = rsConsult!ternom
   apellido = rsConsult!terape
   nombre2 = rsConsult!ternom2
   apellido2 = rsConsult!terape2
   Legajo = rsConsult!empleg
   
'   If IsNull(rsConsult!empremu) Then
'      sueldo = 0
'   Else
'      sueldo = rsConsult!empremu
'   End If
Else
   Flog.writeline "Error al obtener los datos del empleado"
   GoTo MError
End If


'------------------------------------------------------------------
'Busco el valor del cuil
'------------------------------------------------------------------
StrSql = " SELECT cuil.nrodoc "
StrSql = StrSql & " FROM tercero LEFT JOIN ter_doc cuil ON (tercero.ternro=cuil.ternro and cuil.tidnro=10) "
StrSql = StrSql & " WHERE tercero.ternro= " & ternro
       
OpenRecordset StrSql, rsConsult
Cuil = ""
If Not rsConsult.EOF Then
    If Not IsNull(rsConsult!nrodoc) Then
        Cuil = rsConsult!nrodoc
    End If
Else
    Flog.writeline "Error al obtener los datos del cuil"
'   GoTo MError
End If


'------------------------------------------------------------------
'Busco los datos del periodo actual
'------------------------------------------------------------------
StrSql = " SELECT periodo.* FROM periodo "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " AND proceso.pronro= " & pronro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   pliqnro = rsConsult!pliqnro
   pliqmes = rsConsult!pliqmes
   pliqanio = rsConsult!pliqanio
   pliqdesc = rsConsult!pliqdesc
Else
   Flog.writeline "Error al obtener los datos del periodo actual"
   GoTo MError
End If

'------------------------------------------------------------------
'Busco el valor del sueldo basico
'------------------------------------------------------------------
'si el valor sueldo es cero en los datos del empleado entonces tengo que
'buscar el valor del sueldo

'If sueldo = 0 Then
    StrSql = " SELECT almonto"
    StrSql = StrSql & " From acu_liq"
    StrSql = StrSql & " Where acunro = " & acunroSueldo
    StrSql = StrSql & " AND cliqnro = " & cliqnro

    OpenRecordset StrSql, rsConsult

    If Not rsConsult.EOF Then
       sueldo = rsConsult!almonto
    Else
       Flog.writeline "Error al obtener los datos del sueldo"
       sueldo = 0
       'GoTo MError
    End If
    
    rsConsult.Close
'End If

' -------------------------------------------------------------------------
' Busco los datos de la empresa
'--------------------------------------------------------------------------

EmpEstrnro = 0
EmpTernro = 0
StrSql = "SELECT empresa.ternro, empresa.empnom, empresa.estrnro " & _
    " From his_estructura" & _
    " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro" & _
    " WHERE his_estructura.htetdesde <=" & ConvFecha(profecpago) & " AND " & _
    "(his_estructura.htethasta >= " & ConvFecha(profecpago) & " OR his_estructura.htethasta IS NULL)" & _
    "AND his_estructura.ternro = " & ternro & _
    "AND his_estructura.tenro  = 10"

OpenRecordset StrSql, rsConsult

If rsConsult.EOF Then
    Flog.writeline "No se encontró la empresa"
    GoTo MError
Else
    EmpTernro = rsConsult!ternro
    EmpEstrnro = rsConsult!estrnro
    EmpNombre = rsConsult!empnom
End If

rsConsult.Close


' -------------------------------------------------------------------------
'Consulta para obtener la direccion de la empresa
' -------------------------------------------------------------------------

EmpDire = ""
StrSql = "SELECT detdom.calle,detdom.nro,localidad.locdesc,piso,codigopostal From cabdom " & _
    " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & EmpTernro & _
    " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
    Flog.writeline "No se encontró el domicilio de la empresa"
    'Exit Sub
Else
    EmpDire = rsConsult!calle & " " & rsConsult!nro & " Piso " & rsConsult!piso & "<br>" & rsConsult!codigopostal & " - " & rsConsult!locdesc
End If

rsConsult.Close


' -------------------------------------------------------------------------
'Consulta para obtener el cuit de la empresa
' -------------------------------------------------------------------------

EmpCuit = "  "
StrSql = "SELECT cuit.nrodoc FROM tercero " & _
         " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro and cuit.tidnro = 6)" & _
         " Where tercero.ternro =" & EmpTernro
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
    Flog.writeline "No se encontró el CUIT de la Empresa"
    'Exit Sub
Else
    EmpCuit = rsConsult!nrodoc
End If

rsConsult.Close


' -------------------------------------------------------------------------
'Consulta para buscar el logo de la empresa
' -------------------------------------------------------------------------

EmpLogo = ""
EmpLogoAlto = 0
EmpLogoAncho = 0
StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
    " From ter_imag " & _
    " INNER JOIN tipoimag ON tipoimag.tipimnro = 1 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
    " AND ter_imag.ternro =" & EmpTernro
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
    Flog.writeline "No se encontró el Logo de la Empresa"
    'Exit Sub
Else
    EmpLogo = rsConsult!tipimdire & rsConsult!terimnombre
    EmpLogoAlto = rsConsult!tipimaltodef
    EmpLogoAncho = rsConsult!tipimanchodef
End If

rsConsult.Close


' -------------------------------------------------------------------------
'Consulta para buscar la firma de la empresa
' -------------------------------------------------------------------------

EmpFirma = ""
EmpFirmaAlto = 0
EmpFirmaAncho = 0
StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
    " From ter_imag " & _
    " INNER JOIN tipoimag ON tipoimag.tipimnro = 2 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
    " AND ter_imag.ternro =" & EmpTernro
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
    Flog.writeline "No se encontró el Firma de la Empresa"
    'Exit Sub
Else
    EmpFirma = rsConsult!tipimdire & rsConsult!terimnombre
    EmpFirmaAlto = rsConsult!tipimaltodef
    EmpFirmaAncho = rsConsult!tipimanchodef
End If

rsConsult.Close


'------------------------------------------------------------------
'Busco el valor de la sucursal
'------------------------------------------------------------------

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(profecpago) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(profecpago) & ") And his_estructura.tenro = 1 And his_estructura.ternro = " & ternro
       
OpenRecordset StrSql, rsConsult

Sucursal = ""

If Not rsConsult.EOF Then
   Sucursal = rsConsult!estrdabr
Else
   Flog.writeline "Error al obtener los datos de la sucursal"
End If

rsConsult.Close

'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------

StrSql = " INSERT INTO rep_07 "

StrSql = StrSql & "(bpronro , ternro, pronro, apellido,"
StrSql = StrSql & " nombre , apellido2, nombre2, Legajo,"
StrSql = StrSql & " pliqnro , pliqdesc, neto, profecpago,"
StrSql = StrSql & " EmpNombre, descripcion,"
StrSql = StrSql & " orden, EmpDire, EmpCuit, EmpLogo,EmpLogoAlto,"
StrSql = StrSql & " EmpLogoAncho, EmpFirma, EmpFirmaAlto, EmpFirmaAncho,"
StrSql = StrSql & " pliqmes, pliqanio, Auxchar5, cuil"
StrSql = StrSql & " )"
StrSql = StrSql & " VALUES"

StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & ternro
StrSql = StrSql & "," & pronro
StrSql = StrSql & ",'" & apellido & "'"
StrSql = StrSql & ",'" & nombre & "'"
StrSql = StrSql & ",'" & apellido2 & "'"
StrSql = StrSql & ",'" & nombre2 & "'"
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & ",'" & pliqdesc & "' "
StrSql = StrSql & "," & numberForSQL(sueldo)
StrSql = StrSql & ",'" & profecpago & "'"
StrSql = StrSql & ",'" & EmpNombre & "'"
StrSql = StrSql & ",'" & Mid(tituloReporte, 1, 100) & "'"

StrSql = StrSql & "," & orden
StrSql = StrSql & ",'" & Mid(EmpDire, 1, 100) & "'"
StrSql = StrSql & ",'" & Mid(EmpCuit, 1, 30) & "'"
StrSql = StrSql & ",'" & EmpLogo & "'"
StrSql = StrSql & "," & EmpLogoAlto
StrSql = StrSql & "," & EmpLogoAncho
StrSql = StrSql & ",'" & Mid(EmpFirma, 1, 100) & "'"
StrSql = StrSql & "," & EmpFirmaAlto
StrSql = StrSql & "," & EmpFirmaAncho
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & ",'" & Mid(Sucursal, 1, 100) & "'"
StrSql = StrSql & ",'" & Cuil & "'"
StrSql = StrSql & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

MyBeginTrans

objConn.Execute StrSql, , adExecuteNoRecords

repnro = CLng(getLastIdentity(objConn, "rep_07"))

MyCommitTrans

'------------------------------------------------------------------
'Guardo El detalle del recivo en la BD
'------------------------------------------------------------------

Dim a As Integer
Dim concepto As String
Dim estado As Boolean

'Asumo que 0 es el neto
For a = 1 To cantConAcu - 1
    '0=concnrocol, 1=confestiq, 2=conftipo, 3=confval, 4=confval2
    estado = True
    If ArrConcAcu(a, 2) = "CO" Then
        If ArrConcAcu(a, 4) <> "" Then
            concepto = ArrConcAcu(a, 4)
        Else
            concepto = ArrConcAcu(a, 3)
        End If
        
        StrSql = " SELECT detliq.dlimonto monto "
        StrSql = StrSql & " FROM detliq "
        StrSql = StrSql & " INNER JOIN concepto ON detliq.concnro = concepto.concnro "
        StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " "
        'StrSql = StrSql & " And detliq.concnro = " & concepto
        StrSql = StrSql & " And concepto.conccod = " & concepto
        OpenRecordset StrSql, rsDet
        If Not rsDet.EOF Then
            dlimonto = rsDet!Monto
        Else
            dlimonto = 0
            estado = False
        End If
        rsDet.Close
    ElseIf ArrConcAcu(a, 2) = "AC" Then
        StrSql = " SELECT almonto monto "
        StrSql = StrSql & " From acu_liq "
        StrSql = StrSql & " Where acunro = " & ArrConcAcu(a, 3)
        StrSql = StrSql & " AND cliqnro = " & cliqnro
        OpenRecordset StrSql, rsDet
        'dlimonto = rsDet!Monto
        If Not rsDet.EOF Then
            dlimonto = rsDet!Monto
        Else
            dlimonto = 0
            estado = False
        End If
        rsDet.Close
    Else
        Flog.writeline "No se encontró el Concepto o Acumulador"
        
    End If
    
    If estado Then
        StrSql = " INSERT INTO rep_07_det "
        StrSql = StrSql & " (repnro , bpronro, ternro, pronro, concabr, conccod, concnro, dlimonto)"
        StrSql = StrSql & " VALUES"
        StrSql = StrSql & " ("
        StrSql = StrSql & " " & repnro
        StrSql = StrSql & " ," & NroProceso
        StrSql = StrSql & " ," & ternro
        StrSql = StrSql & " ," & pronro
        StrSql = StrSql & " ,'" & ArrConcAcu(a, 1) & "'"
        StrSql = StrSql & " ,'" & ArrConcAcu(a, 2) & "'"
        StrSql = StrSql & " ,'" & ArrConcAcu(a, 3) & "'"
        StrSql = StrSql & " ," & numberForSQL(dlimonto)
        StrSql = StrSql & " )"
        
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
Next a





Exit Sub

MError:
    Flog.writeline "Error en empleado: " & Legajo & " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub

'--------------------------------------------------------------------
' Se encarga de generar un ResultSet de los empleados a cambiar
' si el RS es vacio significa que hay que aplicarlo sobre todos
'--------------------------------------------------------------------
Sub CargarEmpleados(NroProc, ByRef rsEmpl As ADODB.Recordset)

Dim StrEmpl As String

    StrEmpl = "SELECT * FROM batch_empleado WHERE bpronro = " & NroProc
    
    OpenRecordset StrEmpl, rsEmpl
End Sub

Function numberForSQL(Str)
   
  numberForSQL = Replace(Str, ",", ".")

End Function

