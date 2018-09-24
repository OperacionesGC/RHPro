Attribute VB_Name = "RepDotacionPersonal"
'Global Const Version = "1.00" ' Sebastian Stremel
'Global Const FechaModificacion = "14/08/2012"
'Global Const UltimaModificacion = "" 'Version Inicial


'Global Const Version = "1.01" ' Sebastian Stremel
'Global Const FechaModificacion = "30/08/2012"
'Global Const UltimaModificacion = "" 'correccion en consultas

'Global Const Version = "1.02" ' Sebastian Stremel
'Global Const FechaModificacion = "31/08/2012"
'Global Const UltimaModificacion = "" 'se quito modulo de versiones de liquidacion que estaba por erro

'Global Const Version = "1.03" ' Carmen Quintero
'Global Const FechaModificacion = "29/04/2014"
'Global Const UltimaModificacion = "" 'CAS-22837 - VESTIDITOS - Bug en Reporte de ADP
                                     ' - Se modificó la consulta para que los empleados activos se consideren por fase activa
                                     
Global Const Version = "1.04" ' Carmen Quintero
Global Const FechaModificacion = "12/05/2014"
Global Const UltimaModificacion = "" 'CAS-22837 - VESTIDITOS - Bug en Reporte de ADP [Entrega 2]
                                     ' - Se modificaron las consultas del reporte
'--------------------------------------------------------------
'--------------------------------------------------------------
Option Explicit

Dim fs, f
Global NroProceso As Long
Global HuboErrores As Boolean
Global IdUser As String
Global Fecha As Date
Global Hora As String

Global listapronro       'Lista de procesos

Global totalEmpleados
Global cantRegistros
Global CantEmpGrabados As Long 'Cantidad de empleados grabados
Global EmpErrores

Global empresa
Global fechadesde
Global fechahasta
Global legDesde
Global legHasta
Global tenro1
Global estrnro1
Global tenro2
Global estrnro2
Global orden
Global tituloReporte
Global listaSedes
Global listaSectores
Global arrSedes()
Global arrSectores()



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


Dim proNro As Long
Dim ternro  As Long
Dim arrpronro
Dim rsEmpl As New ADODB.Recordset
Dim rsAge As New ADODB.Recordset
Dim rsEmpresas As New ADODB.Recordset
Dim rsPeriodo As New ADODB.Recordset

'Dim acunroSueldo
Dim I
Dim PID As String

Dim parametros As String
Dim ArrParametros
Dim strTempo As String
Dim orden As String

    
Dim arrpliqnro
Dim listapliqnro
Dim pliqNro As Long
Dim pliqMes As Long
Dim pliqAnio As Long
Dim rsConsult2 As New ADODB.Recordset

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
    
    Nombre_Arch = PathFLog & "ReporteDotacionPersonal" & "-" & NroProceso & ".log"
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
    
    Flog.writeline "Inicio Proceso: " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    'OpenConnection strconexion, objConn
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    HuboErrores = False
    
    On Error GoTo CE
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       
       'Obtengo los parametros del proceso
       IdUser = objRs!IdUser
       Fecha = objRs!bprcfecha
       Hora = objRs!bprchora
       parametros = objRs!bprcparam
       Flog.writeline " parametros del proceso --> " & parametros
       ArrParametros = Split(parametros, "@")
       Flog.writeline " limite del array --> " & UBound(ArrParametros)
       
              
       'EMPIEZA EL PROCESO
       Flog.writeline "Generando el reporte"
                    
                  
       'Obtengo los empleados sobre los que tengo que generar el reporte
       cargarEmpleados NroProceso, ArrParametros, rsEmpl
       
       'If Not rsEmpl.EOF Then
            'cantRegistros = rsEmpl.RecordCount
       '     Flog.writeline "Cantidad de empleados a procesar: " & cantRegistros
       '     CantEmpGrabados = 0 'Cantidad de empleados Guardados
       'Else
       '     Flog.writeline "No hay empleados para el filtro seleccionado."
       '     Exit Sub
       'End If
    
       'Actualizo Barch Proceso
       'StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
       '         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
       '         ", bprcempleados ='" & CStr(rsEmpl.RecordCount) & "' WHERE bpronro = " & NroProceso
    
       'objConn.Execute StrSql, , adExecuteNoRecords
 
        'Actualizo
        'StrSql = "UPDATE batch_proceso SET bprcprogreso = 100 " & _
         '        ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
         '        ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
                
        'objConn.Execute StrSql, , adExecuteNoRecords
       
       'Verifico que batch_empleado tenga registros
       'If Not rsEmpl.EOF Then
       '     EmpErrores = False
       '     ternro = rsEmpl!ternro
       '     Flog.writeline ""
       '     Flog.writeline "Generando datos de los empleados "
                    
            'Call ReporteEvolPersonal
                                            
            'Actualizo el estado del proceso
       '     TiempoAcumulado = GetTickCount
            
            'Resto uno a la cantidad de registros
            'cantRegistros = rsEmpl.RecordCount
            

            
            'Borro batch empleado
            '****************************************************************
        '    StrSql = "DELETE  FROM batch_empleado "
        '    StrSql = StrSql & " WHERE bpronro = " & NroProceso
        '    StrSql = StrSql & " AND ternro = " & ternro
        '    objConn.Execute StrSql, , adExecuteNoRecords
            
       'End If
       'rsEmpl.Close
       Set rsEmpl = Nothing
       
       'objRs.Close
       Set objRs = Nothing
    
    Else
        'objRs.Close
        Set objRs = Nothing
        
        objConn.Close
        Set objConn = Nothing
        
        Exit Sub
    End If
       
    
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
       GoTo CE
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline
    Flog.writeline "************************************************************"
    Flog.writeline "Fin :" & Now
    'Flog.writeline "Cantidad de empleados guardados en el reporte: " & cantRegistros
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now
    Flog.writeline " Ultimo SQL: " & StrSql
End Sub

Sub cargarEmpleados(NroProceso, ArrParametros, rsEmpl)


Dim I
Dim j
Dim k
Dim cantidadFecha1
Dim cantidadFecha2
Dim variacion
Dim nombreEmpresa
Dim descripcionReporte
Dim porcVar
Dim nroOrden
Dim porcentajeTotal As Double
Dim porcentaje1 As Double
Dim porcentaje2 As Double

On Error GoTo ERROR

porcentaje1 = 0
porcentaje2 = 0
porcentajeTotal = 0
       'obtengo la empresa
       If CLng(ArrParametros(1)) <> 0 Then
            empresa = CLng(ArrParametros(1))
            Flog.writeline "La codigo de la empresa elegida es: " & ArrParametros(1)
       End If
       
       'Obtengo la fecha desde
       fechadesde = ArrParametros(2)
       Flog.writeline "Se selecciono el parametro fecha desde. " & ArrParametros(2)
       If Len(fechadesde) = 0 Then
            Flog.writeline "No Se selecciono el parametro fecha desde. "
            HuboErrores = True
       End If

       'Obtengo la fecha hasta
       fechahasta = ArrParametros(3)
       Flog.writeline "Se selecciono el parametro fecha hasta. " & ArrParametros(3)
       If Len(fechahasta) = 0 Then
            Flog.writeline "No Se selecciono el parametro fecha hasta. "
            HuboErrores = True
       End If
       
       'Obtengo el tipo de estructura 1 si se configuró
       If CLng(ArrParametros(4)) <> 0 Then
            tenro1 = CLng(ArrParametros(4))
            Flog.writeline "Se selecciono el parametro Tipo de Estructura 1. " & ArrParametros(4)
       End If
       
       'Obtengo la estructura 1 si se configuró
       If CLng(ArrParametros(5)) <> 0 Then
            estrnro1 = CLng(ArrParametros(5))
            Flog.writeline "Se selecciono el parametro Estructura 1. " & ArrParametros(5)
       Else
            estrnro1 = 0
       End If
       
       'Obtengo el tipo de estructura 2 si se configuró
       If CLng(ArrParametros(6)) <> 0 Then
            tenro2 = CLng(ArrParametros(6))
            Flog.writeline "Se selecciono el parametro Tipo de Estructura 2. " & ArrParametros(6)
       End If
       
       'Obtengo la estructura 2 si se configuró
       If CLng(ArrParametros(7)) <> 0 Then
            estrnro2 = CLng(ArrParametros(7))
            Flog.writeline "Se selecciono el parametro Estructura 2. " & ArrParametros(7)
       Else
            estrnro2 = 0
       End If
       
       
       'Obtengo el orden
       orden = ArrParametros(8)
       Flog.writeline "Se selecciono el parametro orden. " & ArrParametros(8)
       
       tituloReporte = ArrParametros(9)
       Flog.writeline "Se selecciono el parametro titulo del reporte. " & ArrParametros(9)

'BUSCO EL NOMBRE DE LA EMPRESA
StrSql = "SELECT estrdabr FROM estructura WHERE estrnro=" & empresa
OpenRecordset StrSql, objRs
If Not objRs.EOF Then
    nombreEmpresa = objRs!estrdabr
    Flog.writeline "Nombre Empresa: " & nombreEmpresa
Else
    Flog.writeline "No se encontro el nombre de la empresa"
End If

descripcionReporte = "Empresa:  " & nombreEmpresa & " Fecha1: " & fechadesde & " Fecha2: " & fechahasta
'INSERTO EN LA CABECERA DEL REPORTE
StrSql = "INSERT INTO rep_dot_per "
StrSql = StrSql & "(bpronro,repdesabr,repdesext,repfecdesde,repfechasta,tenro1,estrnro1,tenro2,estrnro2)"
StrSql = StrSql & " VALUES "
StrSql = StrSql & "( "
StrSql = StrSql & NroProceso
StrSql = StrSql & ", '" & descripcionReporte & "' "
StrSql = StrSql & ", '" & descripcionReporte & "' "
StrSql = StrSql & ", " & ConvFecha(fechadesde)
StrSql = StrSql & ", " & ConvFecha(fechahasta)
StrSql = StrSql & ", " & tenro1
StrSql = StrSql & ", " & estrnro1
StrSql = StrSql & ", " & tenro2
StrSql = StrSql & ", " & estrnro2
StrSql = StrSql & ")"
objConn.Execute StrSql, , adExecuteNoRecords

I = 0
If estrnro1 = 0 Then ' todas las sedes
    StrSql = " SELECT * FROM estructura "
    StrSql = StrSql & " WHERE tenro=" & tenro1
    StrSql = StrSql & " ORDER BY estrdabr " & orden
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Do While Not objRs.EOF
            ReDim Preserve arrSedes(I)
            arrSedes(I) = objRs!estrnro
            I = I + 1
        objRs.MoveNext
        Loop
    'porcentaje1 = 99 / UBound(arrSedes)
    'porcentaje1 = FormatNumber(porcentaje1, 2)
    porcentaje1 = objRs.RecordCount
    End If
    objRs.Close
Else
    porcentaje1 = 1

End If



nroOrden = 0
I = 0
If estrnro2 = 0 Then ' todos los sectores
    StrSql = " SELECT * FROM estructura "
    StrSql = StrSql & " WHERE tenro=" & tenro2
    StrSql = StrSql & " ORDER BY estrdabr " & orden
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Do While Not objRs.EOF
            ReDim Preserve arrSectores(I)
            arrSectores(I) = objRs!estrnro
            I = I + 1
        objRs.MoveNext
        Loop
    porcentaje2 = 99 / objRs.RecordCount
    'porcentaje2 = (45 / UBound(arrSectores)) / UBound(arrSedes)
    porcentaje2 = FormatNumber(porcentaje2, 2)
    porcentaje2 = porcentaje2 / porcentaje1
    End If
    objRs.Close
Else
    porcentaje2 = 1
    porcentaje2 = FormatNumber(porcentaje2, 2)
    porcentaje2 = porcentaje2 / porcentaje1
End If



'para cada estructura, para cada sector cuento los empleados
If estrnro1 = 0 Then ' si eligio todas las sedes
    For j = 0 To UBound(arrSedes)
        If estrnro2 = 0 Then ' si eligio todos los sectores
            For k = 0 To UBound(arrSectores)
                'busco para la fecha 1
                
                
                StrSql = "SELECT COUNT(DISTINCT EMPLEG) cant1 FROM empleado "
                StrSql = StrSql & " INNER JOIN his_estructura empresa on empresa.estrnro =" & empresa & " AND empresa.ternro=empleado.ternro AND empresa.htetdesde <=" & ConvFecha(fechadesde) & " AND  (empresa.htethasta >= " & ConvFecha(fechadesde) & " or empresa.htethasta is null)"
                StrSql = StrSql & " INNER JOIN his_estructura sede on sede.estrnro =" & arrSedes(j) & " AND sede.ternro=empleado.ternro AND sede.htetdesde <=" & ConvFecha(fechadesde) & " AND  (sede.htethasta >= " & ConvFecha(fechadesde) & " or sede.htethasta is null)"
                StrSql = StrSql & " INNER JOIN his_estructura sector on sector.estrnro =" & arrSectores(k) & " AND sector.ternro=empleado.ternro AND sector.htetdesde <=" & ConvFecha(fechadesde) & " AND (sector.htethasta >= " & ConvFecha(fechadesde) & " or sector.htethasta is null)"
                'Comentado el 29/04/2014
                'StrSql = StrSql & " WHERE empleado.empest=-1"
                StrSql = StrSql & " WHERE (empleado.ternro in (SELECT empleado from fases WHERE fases.empleado = empleado.ternro "
                StrSql = StrSql & " AND (fases.altfec <=" & ConvFecha(fechadesde) & " AND (fases.bajfec is null or fases.bajfec>=" & ConvFecha(fechadesde) & ")) )) "
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    cantidadFecha1 = objRs!cant1
                    objRs.Close
                Else
                    cantidadFecha1 = 0
                    objRs.Close
                End If
                
                
                'busco para la fecha2
                StrSql = "SELECT COUNT(DISTINCT EMPLEG) cant2 FROM empleado "
                StrSql = StrSql & " INNER JOIN his_estructura empresa on empresa.estrnro =" & empresa & " AND empresa.ternro=empleado.ternro AND empresa.htetdesde <=" & ConvFecha(fechahasta) & " AND (empresa.htethasta >= " & ConvFecha(fechahasta) & " or empresa.htethasta is null)"
                StrSql = StrSql & " INNER JOIN his_estructura sede on sede.estrnro =" & arrSedes(j) & " AND sede.ternro=empleado.ternro AND sede.htetdesde <=" & ConvFecha(fechahasta) & " AND (sede.htethasta >= " & ConvFecha(fechahasta) & " or sede.htethasta is null)"
                StrSql = StrSql & " INNER JOIN his_estructura sector on sector.estrnro =" & arrSectores(k) & " AND sector.ternro=empleado.ternro AND sector.htetdesde <=" & ConvFecha(fechahasta) & " AND (sector.htethasta >= " & ConvFecha(fechahasta) & " or sector.htethasta is null)"
                'Comentado el 29/04/2014
                'StrSql = StrSql & " WHERE empleado.empest=-1"
                StrSql = StrSql & " WHERE (empleado.ternro in (SELECT empleado from fases WHERE fases.empleado = empleado.ternro "
                StrSql = StrSql & " AND (fases.altfec <=" & ConvFecha(fechahasta) & " AND (fases.bajfec is null or fases.bajfec>=" & ConvFecha(fechahasta) & ")) )) "
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    cantidadFecha2 = objRs!cant2
                    objRs.Close
                Else
                    cantidadFecha2 = 0
                    objRs.Close
                End If
                
                'CALCULO LA VARIACION
                variacion = cantidadFecha2 - cantidadFecha1
                
                'CALCULO EL % DE VARIACION
                If cantidadFecha1 <> 0 Then
                    porcVar = (variacion / cantidadFecha1) * 100
                Else
                    porcVar = (variacion / 1) * 100
                    'porcVar = 0
                    Flog.writeline "No se puede calcular el porcentaje de variacion porque la cantidad en la fecha1 es 0"
                End If
                'ACA HACER INSERT
                StrSql = " INSERT INTO rep_dot_per_det "
                StrSql = StrSql & " (bpronro,reporden,tenro1,estrnro1,tenro2,estrnro2,repdetcant1,repdetcant2,repdetvar,repdetpor) "
                StrSql = StrSql & " VALUES "
                StrSql = StrSql & "("
                StrSql = StrSql & NroProceso
                StrSql = StrSql & ", " & nroOrden
                StrSql = StrSql & ", 49"
                StrSql = StrSql & ", " & arrSedes(j)
                StrSql = StrSql & ", 2"
                StrSql = StrSql & ", " & arrSectores(k)
                StrSql = StrSql & ", " & cantidadFecha1
                StrSql = StrSql & ", " & cantidadFecha2
                StrSql = StrSql & ", " & variacion
                StrSql = StrSql & ", " & porcVar
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                'HASTA ACA
                nroOrden = nroOrden + 1
                'porcentajeTotal = porcentajeTotal + (porcentaje1 + porcentaje2)
                porcentajeTotal = porcentajeTotal + porcentaje2
                'porcentajeTotal = CDbl(porcentajeTotal)
                StrSql = "UPDATE batch_proceso SET  bprcprogreso =" & porcentajeTotal & " , bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesando' WHERE bpronro = " & NroProceso
                objConn.Execute StrSql, , adExecuteNoRecords
            Next
        Else
                'busco para la fecha 1
                StrSql = "SELECT COUNT(DISTINCT EMPLEG) cant1 FROM empleado "
                StrSql = StrSql & " INNER JOIN his_estructura empresa on empresa.estrnro =" & empresa & " AND empresa.ternro=empleado.ternro AND empresa.htetdesde <=" & ConvFecha(fechadesde) & " AND  (empresa.htethasta >= " & ConvFecha(fechadesde) & " or empresa.htethasta is null)"
                StrSql = StrSql & " INNER JOIN his_estructura sede on sede.estrnro =" & arrSedes(j) & " AND sede.ternro=empleado.ternro AND sede.htetdesde <=" & ConvFecha(fechadesde) & " AND  (sede.htethasta >= " & ConvFecha(fechadesde) & " or sede.htethasta is null)"
                StrSql = StrSql & " INNER JOIN his_estructura sector on sector.estrnro =" & estrnro2 & " AND sector.ternro=empleado.ternro AND sector.htetdesde <=" & ConvFecha(fechadesde) & " AND (sector.htethasta >= " & ConvFecha(fechadesde) & " or sector.htethasta is null)"
                'Comentado el 29/04/2014
                'StrSql = StrSql & " WHERE empleado.empest=-1"
                StrSql = StrSql & " WHERE (empleado.ternro in (SELECT empleado from fases WHERE fases.empleado = empleado.ternro "
                StrSql = StrSql & " AND (fases.altfec <=" & ConvFecha(fechadesde) & " AND (fases.bajfec is null or fases.bajfec>=" & ConvFecha(fechadesde) & ")) )) "
                
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    cantidadFecha1 = objRs!cant1
                    objRs.Close
                Else
                    cantidadFecha1 = 0
                    objRs.Close
                End If
                
                'busco para la fecha2
                StrSql = "SELECT COUNT(DISTINCT EMPLEG) cant2 FROM empleado "
                StrSql = StrSql & " INNER JOIN his_estructura empresa on empresa.estrnro =" & empresa & " AND empresa.ternro=empleado.ternro AND empresa.htetdesde <=" & ConvFecha(fechahasta) & " AND (empresa.htethasta >= " & ConvFecha(fechahasta) & " or empresa.htethasta is null)"
                StrSql = StrSql & " INNER JOIN his_estructura sede on sede.estrnro =" & arrSedes(j) & " AND sede.ternro=empleado.ternro AND sede.htetdesde <=" & ConvFecha(fechahasta) & " AND (sede.htethasta >= " & ConvFecha(fechahasta) & " or sede.htethasta is null)"
                StrSql = StrSql & " INNER JOIN his_estructura sector on sector.estrnro =" & estrnro2 & " AND sector.ternro=empleado.ternro AND sector.htetdesde <=" & ConvFecha(fechahasta) & " AND (sector.htethasta >= " & ConvFecha(fechahasta) & " or sector.htethasta is null)"
                'Comentado el 29/04/2014
                'StrSql = StrSql & " WHERE empleado.empest=-1"
                StrSql = StrSql & " WHERE (empleado.ternro in (SELECT empleado from fases WHERE fases.empleado = empleado.ternro "
                StrSql = StrSql & " AND (fases.altfec <=" & ConvFecha(fechahasta) & " AND (fases.bajfec is null or fases.bajfec>=" & ConvFecha(fechahasta) & ")) )) "
                
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    cantidadFecha2 = objRs!cant2
                    objRs.Close
                Else
                    cantidadFecha2 = 0
                    objRs.Close
                End If

                'CALCULO LA VARIACION
                variacion = cantidadFecha2 - cantidadFecha1
                
                'CALCULO EL % DE VARIACION
                If cantidadFecha1 <> 0 Then
                    porcVar = (variacion / cantidadFecha1) * 100
                Else
                    porcVar = (variacion / 1) * 100
                    Flog.writeline "No se puede calcular el porcentaje de variacion porque la cantidad en la fecha1 es 0"
                End If
                'ACA HACER INSERT
                StrSql = " INSERT INTO rep_dot_per_det "
                StrSql = StrSql & " (bpronro,reporden,tenro1,estrnro1,tenro2,estrnro2,repdetcant1,repdetcant2,repdetvar,repdetpor) "
                StrSql = StrSql & " VALUES "
                StrSql = StrSql & "("
                StrSql = StrSql & NroProceso
                StrSql = StrSql & ", " & nroOrden
                StrSql = StrSql & ", 49"
                StrSql = StrSql & ", " & arrSedes(j)
                StrSql = StrSql & ", 2"
                StrSql = StrSql & ", " & estrnro2
                StrSql = StrSql & ", " & cantidadFecha1
                StrSql = StrSql & ", " & cantidadFecha2
                StrSql = StrSql & ", " & variacion
                StrSql = StrSql & ", " & porcVar
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                'HASTA ACA
                nroOrden = nroOrden + 1
                'porcentajeTotal = porcentajeTotal + (porcentaje1 + porcentaje2)
                porcentajeTotal = porcentajeTotal + porcentaje2
                StrSql = "UPDATE batch_proceso SET  bprcprogreso =" & porcentajeTotal & " , bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesando' WHERE bpronro = " & NroProceso
                objConn.Execute StrSql, , adExecuteNoRecords

        End If

    Next
Else ' SI ELIGIO UNA SEDE SOLA
    If estrnro2 = 0 Then ' si eligio todos los sectores
        For k = 0 To UBound(arrSectores)
            'busco para la fecha 1
            StrSql = "SELECT COUNT(DISTINCT EMPLEG) cant1 FROM empleado "
            StrSql = StrSql & " INNER JOIN his_estructura empresa on empresa.estrnro =" & empresa & " AND empresa.ternro=empleado.ternro AND empresa.htetdesde <=" & ConvFecha(fechadesde) & " AND  (empresa.htethasta >= " & ConvFecha(fechadesde) & " or empresa.htethasta is null)"
            StrSql = StrSql & " INNER JOIN his_estructura sede on sede.estrnro =" & estrnro1 & " AND sede.ternro=empleado.ternro AND sede.htetdesde <=" & ConvFecha(fechadesde) & " AND  (sede.htethasta >= " & ConvFecha(fechadesde) & " or sede.htethasta is null)"
            StrSql = StrSql & " INNER JOIN his_estructura sector on sector.estrnro =" & arrSectores(k) & " AND sector.ternro=empleado.ternro AND sector.htetdesde <=" & ConvFecha(fechadesde) & " AND (sector.htethasta >= " & ConvFecha(fechadesde) & " or sector.htethasta is null)"
            'Comentado el 29/04/2014
            'StrSql = StrSql & " WHERE empleado.empest=-1"
            StrSql = StrSql & " WHERE (empleado.ternro in (SELECT empleado from fases WHERE fases.empleado = empleado.ternro "
            StrSql = StrSql & " AND (fases.altfec <=" & ConvFecha(fechadesde) & " AND (fases.bajfec is null or fases.bajfec>=" & ConvFecha(fechadesde) & ")) )) "
            OpenRecordset StrSql, objRs
            If Not objRs.EOF Then
                cantidadFecha1 = objRs!cant1
                objRs.Close
            Else
                cantidadFecha1 = 0
                objRs.Close
            End If
            
            'busco para la fecha2
            StrSql = "SELECT COUNT(DISTINCT EMPLEG) cant2 FROM empleado "
            StrSql = StrSql & " INNER JOIN his_estructura empresa on empresa.estrnro =" & empresa & " AND empresa.ternro=empleado.ternro AND empresa.htetdesde <=" & ConvFecha(fechahasta) & " AND (empresa.htethasta >= " & ConvFecha(fechahasta) & " or empresa.htethasta is null)"
            StrSql = StrSql & " INNER JOIN his_estructura sede on sede.estrnro =" & estrnro1 & " AND sede.ternro=empleado.ternro AND sede.htetdesde <=" & ConvFecha(fechahasta) & " AND (sede.htethasta >= " & ConvFecha(fechahasta) & " or sede.htethasta is null)"
            StrSql = StrSql & " INNER JOIN his_estructura sector on sector.estrnro =" & arrSectores(k) & " AND sector.ternro=empleado.ternro AND sector.htetdesde <=" & ConvFecha(fechahasta) & " AND (sector.htethasta >= " & ConvFecha(fechahasta) & " or sector.htethasta is null)"
            'Comentado el 29/04/2014
            'StrSql = StrSql & " WHERE empleado.empest=-1"
            StrSql = StrSql & " WHERE (empleado.ternro in (SELECT empleado from fases WHERE fases.empleado = empleado.ternro "
            StrSql = StrSql & " AND (fases.altfec <=" & ConvFecha(fechahasta) & " AND (fases.bajfec is null or fases.bajfec>=" & ConvFecha(fechahasta) & ")) )) "
            
            OpenRecordset StrSql, objRs
            If Not objRs.EOF Then
                cantidadFecha2 = objRs!cant2
                objRs.Close
            Else
                cantidadFecha2 = 0
                objRs.Close
            End If


            'CALCULO LA VARIACION
            variacion = cantidadFecha2 - cantidadFecha1
            
            'CALCULO EL % DE VARIACION
            If cantidadFecha1 <> 0 Then
                porcVar = (variacion / cantidadFecha1) * 100
            Else
                porcVar = (variacion / 1) * 100
                Flog.writeline "No se puede calcular el porcentaje de variacion porque la cantidad en la fecha1 es 0"
            End If
            'ACA HACER INSERT
            StrSql = " INSERT INTO rep_dot_per_det "
            StrSql = StrSql & " (bpronro,reporden,tenro1,estrnro1,tenro2,estrnro2,repdetcant1,repdetcant2,repdetvar,repdetpor) "
            StrSql = StrSql & " VALUES "
            StrSql = StrSql & "("
            StrSql = StrSql & NroProceso
            StrSql = StrSql & ", " & nroOrden
            StrSql = StrSql & ", 49"
            StrSql = StrSql & ", " & estrnro1
            StrSql = StrSql & ", 2"
            StrSql = StrSql & ", " & arrSectores(k)
            StrSql = StrSql & ", " & cantidadFecha1
            StrSql = StrSql & ", " & cantidadFecha2
            StrSql = StrSql & ", " & variacion
            StrSql = StrSql & ", " & porcVar
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            'HASTA ACA
            nroOrden = nroOrden + 1
            'porcentajeTotal = porcentajeTotal + (porcentaje1 + porcentaje2)
            porcentajeTotal = porcentajeTotal + porcentaje2
            StrSql = "UPDATE batch_proceso SET  bprcprogreso =" & porcentajeTotal & " , bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesando' WHERE bpronro = " & NroProceso
            objConn.Execute StrSql, , adExecuteNoRecords

        Next
    Else
            'busco para la fecha 1
            StrSql = "SELECT COUNT(DISTINCT EMPLEG) cant1 FROM empleado "
            StrSql = StrSql & " INNER JOIN his_estructura empresa on empresa.estrnro =" & empresa & " AND empresa.ternro=empleado.ternro AND empresa.htetdesde <=" & ConvFecha(fechadesde) & " AND  (empresa.htethasta >= " & ConvFecha(fechadesde) & " or empresa.htethasta is null)"
            StrSql = StrSql & " INNER JOIN his_estructura sede on sede.estrnro =" & estrnro1 & " AND sede.ternro=empleado.ternro AND sede.htetdesde <=" & ConvFecha(fechadesde) & " AND  (sede.htethasta >= " & ConvFecha(fechadesde) & " or sede.htethasta is null)"
            StrSql = StrSql & " INNER JOIN his_estructura sector on sector.estrnro =" & estrnro2 & " AND sector.ternro=empleado.ternro AND sector.htetdesde <=" & ConvFecha(fechadesde) & " AND (sector.htethasta >= " & ConvFecha(fechadesde) & " or sector.htethasta is null)"
            'Comentado el 29/04/2014
            'StrSql = StrSql & " WHERE empleado.empest=-1"
            StrSql = StrSql & " WHERE (empleado.ternro in (SELECT empleado from fases WHERE fases.empleado = empleado.ternro "
            StrSql = StrSql & " AND (fases.altfec <=" & ConvFecha(fechadesde) & " AND (fases.bajfec is null or fases.bajfec>=" & ConvFecha(fechadesde) & ")) )) "
            OpenRecordset StrSql, objRs
            If Not objRs.EOF Then
                cantidadFecha1 = objRs!cant1
                objRs.Close
            Else
                cantidadFecha1 = 0
                objRs.Close
            End If
            
            'busco para la fecha2
            StrSql = "SELECT COUNT(DISTINCT EMPLEG) cant2 FROM empleado "
            StrSql = StrSql & " INNER JOIN his_estructura empresa on empresa.estrnro =" & empresa & " AND empresa.ternro=empleado.ternro AND empresa.htetdesde <=" & ConvFecha(fechahasta) & " AND (empresa.htethasta >= " & ConvFecha(fechahasta) & " or empresa.htethasta is null)"
            StrSql = StrSql & " INNER JOIN his_estructura sede on sede.estrnro =" & estrnro1 & " AND sede.ternro=empleado.ternro AND sede.htetdesde <=" & ConvFecha(fechahasta) & " AND (sede.htethasta >= " & ConvFecha(fechahasta) & " or sede.htethasta is null)"
            StrSql = StrSql & " INNER JOIN his_estructura sector on sector.estrnro =" & estrnro2 & " AND sector.ternro=empleado.ternro AND sector.htetdesde <=" & ConvFecha(fechahasta) & " AND (sector.htethasta >= " & ConvFecha(fechahasta) & " or sector.htethasta is null)"
            'Comentado el 29/04/2014
            'StrSql = StrSql & " WHERE empleado.empest=-1"
            StrSql = StrSql & " WHERE (empleado.ternro in (SELECT empleado from fases WHERE fases.empleado = empleado.ternro "
            StrSql = StrSql & " AND (fases.altfec <=" & ConvFecha(fechahasta) & " AND (fases.bajfec is null or fases.bajfec>=" & ConvFecha(fechahasta) & ")) )) "
            
            OpenRecordset StrSql, objRs
            If Not objRs.EOF Then
                cantidadFecha2 = objRs!cant2
                objRs.Close
            Else
                cantidadFecha2 = 0
                objRs.Close
            End If
            
            'CALCULO LA VARIACION
            variacion = cantidadFecha2 - cantidadFecha1
            
            'CALCULO EL % DE VARIACION
            If cantidadFecha1 <> 0 Then
                porcVar = (variacion / cantidadFecha1) * 100
            Else
                porcVar = (variacion / 1) * 100
                Flog.writeline "No se puede calcular el porcentaje de variacion porque la cantidad en la fecha1 es 0"
            End If
            
            'ACA HACER INSERT
            StrSql = " INSERT INTO rep_dot_per_det "
            StrSql = StrSql & " (bpronro,reporden,tenro1,estrnro1,tenro2,estrnro2,repdetcant1,repdetcant2,repdetvar,repdetpor) "
            StrSql = StrSql & " VALUES "
            StrSql = StrSql & "("
            StrSql = StrSql & NroProceso
            StrSql = StrSql & ", " & nroOrden
            StrSql = StrSql & ", 49"
            StrSql = StrSql & ", " & estrnro1
            StrSql = StrSql & ", 2"
            StrSql = StrSql & ", " & estrnro2
            StrSql = StrSql & ", " & cantidadFecha1
            StrSql = StrSql & ", " & cantidadFecha2
            StrSql = StrSql & ", " & variacion
            StrSql = StrSql & ", " & porcVar
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            'HASTA ACA
            nroOrden = nroOrden + 1
            'porcentajeTotal = porcentajeTotal + (porcentaje1 + porcentaje2)
            porcentajeTotal = porcentajeTotal + porcentaje2
            StrSql = "UPDATE batch_proceso SET  bprcprogreso =" & porcentajeTotal & " , bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesando' WHERE bpronro = " & NroProceso
            objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If
    GoTo datosOK
ERROR:
    HuboErrores = True
    Flog.writeline "Error en cargar empleados: " & Err.Description
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql
    Exit Sub
datosOK:
    Flog.writeline "Se cargo el emplado satifactoriamente. "


End Sub
