Attribute VB_Name = "repFormularioF21"
'Global Const Version = "1.00" ' Carmen Quintero
'Global Const FechaModificacion = "28/09/2011"
'Global Const UltimaModificacion = "" 'Version Inicial

Global Const Version = "1.01" ' Sebastian Stremel
Global Const FechaModificacion = "18/10/2012"
Global Const UltimaModificacion = "Se modifico para que cuente la cantidad de empleados que tienen un concepto o un acumulador" 'Version Inicial

'--------------------------------------------------------------
'--------------------------------------------------------------
Option Explicit

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

Global Pagina As Long
Global tipoModelo As Integer
Global arrTipoConc(1000) As Integer
Global tituloReporte As String

Global tenro1 As Integer
Global estrnro1 As Integer
Global tenro2 As Integer
Global estrnro2 As Integer
Global tenro3 As Integer
Global estrnro3 As Integer
Global fecEstr As String

Global empresa As String
Global Empnro As Long
Global Empnroestr As Long
Global Centcostnroestr As Long
Global emprTer As Long
Global emprDire As String
Global emprCuit

Global IdUser As String
Global Fecha As Date
Global Hora As String

Global listapronro       'Lista de procesos

Global totalEmpleados
Global cantRegistros

Global incluyeAgencia As Integer
Global NroAcDiasTrabajados As Long

Global TEAfp As Long
Global TEIsapre As Long
Global acumafp As Long
Global acumisapre As Long


Global VectorEsConc(25) As Boolean 'True = Concepto    False = Acumulador
Global VectorNroACCO(25) As Long
Global VectorValorACCO(25) As Double

'sebastian stremel
Global VectorEsCant(25) As Double
Global VectorEsCantCon(25) As Double
Global VectorEsCantAcu(25) As Double

'hasta aca

Global CantEmpGrabados As Long 'Cantidad de empleados grabados

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

Dim historico As Boolean
'Dim param
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
Dim tituloReporte As String

Dim parametros As String
Dim ArrParametros
Dim strTempo As String
Dim orden As Long
Dim ord
    
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
    
    Nombre_Arch = PathFLog & "ReporteFormularioF21" & "-" & NroProceso & ".log"
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
       
       'Obtengo el numero de empresa
       If CLng(ArrParametros(5)) <> 0 Then
            Empnroestr = CLng(ArrParametros(5))
            Flog.writeline "Se selecciono el parametro Empresa. " & ArrParametros(5)
       Else
            Flog.writeline "No Se selecciono el parametro Empresa. "
            HuboErrores = True
       End If
       
       'Obtengo el numero de centro de costo
       If CLng(ArrParametros(11)) <> 0 Then
            Centcostnroestr = CLng(ArrParametros(11))
            Flog.writeline "Se selecciono el parametro Centro de Costo. " & ArrParametros(11)
       Else
            Flog.writeline "No Se selecciono el parametro Centro de Costo. "
            HuboErrores = True
       End If
       
       'Obtengo el titulo del reporte
       tituloReporte = ArrParametros(13)
       Flog.writeline "Se selecciono el parametro Titulo Reporte. " & ArrParametros(13)
       
       fecEstr = ArrParametros(9)
       Flog.writeline "Se selecciono el parametro fecha estructura " & ArrParametros(9)
        
       'Obtengo la lista de procesos
       listapronro = ArrParametros(4)
       Flog.writeline "Se selecciono el parametro lista de procesos " & ArrParametros(4)
              
       'EMPIEZA EL PROCESO
       Flog.writeline "Generando el reporte"
                
       
       'INICIALIZ0 ARRAY DE CONCEPTOS
        '************************
       For I = 1 To 25
            VectorEsConc(I) = False 'As Boolean
            VectorNroACCO(I) = 0 ' As Long
            VectorValorACCO(I) = 0 'As Double
       Next
       
        TEAfp = 0
        TEIsapre = 0
        
        'Obtengo la Configuracion del Confrep
        StrSql = "SELECT * FROM confrep"
        StrSql = StrSql & " WHERE repnro = 354"
        StrSql = StrSql & " ORDER BY confnrocol"
        OpenRecordset StrSql, objRs2
        
        Flog.writeline "Obtengo los datos del confrep"
        
        If objRs2.EOF Then
          Flog.writeline "No esta configurado el ConfRep para el reporte"
          HuboErrores = True
        End If
       
       
        If Not HuboErrores Then
       
            Do Until objRs2.EOF
                 ' Levanto parametros del Confrep SON 25 CONCEPTOS/ACUMULADORES
               Select Case objRs2!confnrocol
                  Case 1 To 14
                     ' si es concepto
                     If objRs2!conftipo = "CO" Then
                        
                        VectorEsConc(objRs2!confnrocol) = True
                 
                        StrSql = "SELECT concnro FROM concepto WHERE conccod = " & objRs2!confval
                        StrSql = StrSql & " OR conccod = '" & objRs2!confval2 & "'"
                       
                        OpenRecordset StrSql, objRs3
                       
                        If objRs3.EOF Then
                           VectorNroACCO(objRs2!confnrocol) = 0
                        Else
                           VectorNroACCO(objRs2!confnrocol) = CLng(objRs3!concnro)
                        End If
                       
                        objRs3.Close
                     End If
                     ' si es acumulador
                     If objRs2!conftipo = "AC" Then
                        VectorEsConc(objRs2!confnrocol) = False
                        VectorNroACCO(objRs2!confnrocol) = objRs2!confval
                     End If
                     
                     'sebastian nuevo tipo cantidad acumulador
                     If objRs2!conftipo = "CAC" Then
                        VectorEsCantAcu(objRs2!confnrocol) = True
                        VectorNroACCO(objRs2!confnrocol) = objRs2!confval
                     End If
                     
                     
                     'sebastian nuevo tipo cantidad concepto
                     If objRs2!conftipo = "CCO" Then
                        VectorEsCantCon(objRs2!confnrocol) = True
                        VectorNroACCO(objRs2!confnrocol) = objRs2!confval2
                     End If
                     
                  Case 15
                    If objRs2!conftipo = "AC" Then
                        acumafp = objRs2!confval
                    Else
                        Flog.writeline "La columna 15 debe ser el acumulador AFP. "
                    End If
                  
                  Case 16
                    If objRs2!conftipo = "AC" Then
                        acumisapre = objRs2!confval
                    Else
                        Flog.writeline "La columna 16 debe ser el acumulador ISAPRE. "
                    End If
                     
                  Case 43
                    If objRs2!conftipo = "TE" Then
                        TEAfp = objRs2!confval
                    Else
                        Flog.writeline "La columna 43 AFP debe ser un Tipo de Estructura. "
                    End If
                                             
                  Case 44
                    If objRs2!conftipo = "TE" Then
                        TEIsapre = objRs2!confval
                     Else
                        Flog.writeline "La columna 44 ISAPRE debe ser un Tipo de Estructura. "
                    End If
               End Select
            
               objRs2.MoveNext
            Loop

               
           'Obtengo los empleados sobre los que tengo que generar el reporte
           'CargarEmpleados(ByVal NroProc, ByRef rsEmpl As ADODB.Recordset, ByVal empresa As Long)
           CargarEmpleados NroProceso, rsEmpl, 0
           If Not rsEmpl.EOF Then
                'cantRegistros = rsEmpl.RecordCount
                Flog.writeline "Cantidad de empleados a procesar: " & cantRegistros
                CantEmpGrabados = 0 'Cantidad de empleados Guardados
           Else
                Flog.writeline "No hay empleados para el filtro seleccionado."
                Exit Sub
           End If
        
           'Actualizo Barch Proceso
           StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                    ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                    ", bprcempleados ='" & CStr(rsEmpl.RecordCount) & "' WHERE bpronro = " & NroProceso
        
           objConn.Execute StrSql, , adExecuteNoRecords
     
           'Busco los datos de la empresa
           buscarDatosEmpresa (Empnroestr)
              
           '------------------------------------------------------------------
           'Busco El mes y el año correspondiente al periodo
           '------------------------------------------------------------------
           StrSql = " SELECT pliqmes, pliqanio FROM periodo where pliqnro= " & pliqNro
           OpenRecordset StrSql, rsConsult2
        
           If Not rsConsult2.EOF Then
                pliqMes = rsConsult2!pliqMes
                pliqAnio = rsConsult2!pliqAnio
           Else
                pliqMes = 0
                pliqAnio = 0
           End If
        
           rsConsult2.Close
                    
           'Obtengo la lista de procesos
           arrpronro = Split(listapronro, ",")

           ord = 0
     
           'Genero por cada empleado un registro
           If Not rsEmpl.EOF Then
                arrpronro = Split(listapronro, ",")
                EmpErrores = False
                ternro = rsEmpl!ternro
                orden = ord
                Flog.writeline ""
                Flog.writeline "Generando datos de los empleados "
                        
                Call ReporteFormularioF21(arrpronro, ternro, tituloReporte, orden)
                                                
                'Actualizo el estado del proceso
                TiempoAcumulado = GetTickCount
                
                'Resto uno a la cantidad de registros
                cantRegistros = rsEmpl.RecordCount
                
                'Actualizo
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                         ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
                        
                objConn.Execute StrSql, , adExecuteNoRecords
                
                ord = ord + 1
                
                'Borro batch empleado
                '****************************************************************
                StrSql = "DELETE  FROM batch_empleado "
                StrSql = StrSql & " WHERE bpronro = " & NroProceso
                StrSql = StrSql & " AND ternro = " & ternro
                objConn.Execute StrSql, , adExecuteNoRecords
          End If
        End If
    Else
        Exit Sub
    End If
   
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline
    Flog.writeline "************************************************************"
    Flog.writeline "Fin :" & Now
    Flog.writeline "Cantidad de empleados guardados en el reporte: " & cantRegistros
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now
    Flog.writeline " Ultimo SQL: " & StrSql
End Sub


Function controlNull(Str)
  If Trim(Str) = "" Then
     controlNull = "null"
  Else
     controlNull = "'" & Str & "'"
  End If
End Function


'--------------------------------------------------------------------
' Se encarga de generar los datos
'--------------------------------------------------------------------
Sub ReporteFormularioF21(ListaPro, ternro As Long, descripcion As String, orden As Long)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset

Dim rsEstrnro As New ADODB.Recordset
Dim rsConsult2 As New ADODB.Recordset
Dim rsF21 As New ADODB.Recordset

Dim sqlAux As String
Dim rsCant As New ADODB.Recordset


'Variables donde se guardan los datos del INSERT final

Dim Legajo As Long
Dim apellido As String
Dim apellido2 As String
Dim nombre As String
Dim nombre2 As String
Dim pliqNro As Long
Dim pliqMes As Integer
Dim pliqAnio As Long
Dim documento  As String
Dim cliqnro As Long
Dim EmpTernro As Long
Dim DiasTrabajados
Dim estrnomb1
Dim estrnomb2
Dim estrnomb3
Dim tenomb1
Dim tenomb2
Dim tenomb3
Dim proNro
Dim DescEstructura As String
Dim sql As String
Dim I As Integer
Dim direccion As String
Dim G
Dim GrabaEmpleado As Boolean
Dim RUT As String
Dim prodesc As String
Dim contempl As Double
Dim l_monto As Double
Dim l_monto_total As Double
Dim l_ternroanterior As Long
Dim cantEmpl As Long



On Error GoTo MError

'Inicializo Conceptos
For I = 1 To 25
   VectorValorACCO(I) = 0 'As Double
Next

contempl = 0
l_monto = 0

'*********************************************************************
'Ciclo por todos los procesos seleccionados del periodo
'*********************************************************************
GrabaEmpleado = False

For I = 0 To UBound(ListaPro)
                          
   proNro = ListaPro(I)


        '------------------------------------------------------------------
        'Busco los datos del periodo actual
        '------------------------------------------------------------------
        StrSql = " SELECT periodo.*, proceso.profecpago, proceso.prodesc, cabliq.cliqnro FROM periodo "
        StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
        StrSql = StrSql & " AND proceso.pronro= " & proNro
        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
        StrSql = StrSql & " AND cabliq.empleado= " & ternro
        
        '---LOG---
        Flog.writeline "Buscando datos del periodo para el proceso: " & proNro
        
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           pliqNro = rsConsult!pliqNro
           pliqMes = rsConsult!pliqMes
           pliqAnio = rsConsult!pliqAnio
           prodesc = rsConsult!prodesc
           'cliqnro = rsConsult!cliqnro
           'pliqdesde = rsConsult!pliqdesde
           'pliqhasta = rsConsult!pliqhasta
        
           rsConsult.Close
        
            '------------------------------------------------------------------
            'Busco los datos de los 25 acumuladores/conceptos
            '------------------------------------------------------------------
            'VectorEsConc(25) As Boolean 'True = Concepto    False = Acumulador
            'VectorNroACCO(25) As Long  Numero de acumulador o concepto
            'VectorValorACCO(25) As Double
            
            
           For G = 1 To 25
                If VectorNroACCO(G) <> 0 Then 'Si es igual a cero es porq no se configuro valor para ese AC/CO
                    If VectorEsCantAcu(G) = True Then
                        'sebastian stremel 26/06/2012 cuento la cantidad de empleados que tiene el acumulador
                        sqlAux = "SELECT COUNT(distinct empleado) cantidad"
                        sqlAux = sqlAux & " FROM acu_liq"
                        sqlAux = sqlAux & " INNER JOIN cabliq ON cabliq.cliqnro = acu_liq.cliqnro"
                        sqlAux = sqlAux & " WHERE acunro = " & VectorNroACCO(G)
                        sqlAux = sqlAux & " AND proNro = " & proNro
                        OpenRecordset sqlAux, rsCant
                        Flog.writeline "consulta cantidad acumulador: " & sqlAux
                        If Not rsCant.EOF Then
                            VectorValorACCO(G) = rsCant!Cantidad
                        End If
                        rsCant.Close
                        'hasta aca
                    Else
                        If VectorEsCantCon(G) = True Then
                            'sebastian stremel 26/06/2012 cuento la cantidad de empleados que tiene el concepto
                            sql = " SELECT SUM (DISTINCT detliq.dlimonto) monto "
                            sql = sql & " FROM cabliq "
                            sql = sql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro AND cabliq.pronro = " & proNro
                            sql = sql & " INNER JOIN periodo ON proceso.pliqnro = periodo.pliqnro "
                            sql = sql & " INNER JOIN detliq  ON cabliq.cliqnro = detliq.cliqnro  "
                            sql = sql & " INNER JOIN batch_empleado  ON batch_empleado.ternro = cabliq.empleado "
                            sql = sql & " AND detliq.concnro = " & VectorNroACCO(G)
                            OpenRecordset sqlAux, rsCant
                            Flog.writeline "consulta cantidad concepto: " & sqlAux
                            If Not rsCant.EOF Then
                                VectorValorACCO(G) = rsCant!Cantidad
                            End If
                            rsCant.Close
                        'hasta aca
                        Else
                            If VectorEsConc(G) = False Then
                                sql = "SELECT sum(almonto) monto"
                                sql = sql & " FROM acu_liq"
                                sql = sql & " INNER JOIN cabliq ON cabliq.cliqnro = acu_liq.cliqnro"
                                sql = sql & " WHERE acunro = " & VectorNroACCO(G)
                                sql = sql & " AND proNro = " & proNro
                            Else
                                sql = " SELECT SUM (DISTINCT detliq.dlimonto) monto "
                                sql = sql & " FROM cabliq "
                                sql = sql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro AND cabliq.pronro = " & proNro
                                sql = sql & " INNER JOIN periodo ON proceso.pliqnro = periodo.pliqnro "
                                sql = sql & " INNER JOIN detliq  ON cabliq.cliqnro = detliq.cliqnro  "
                                sql = sql & " INNER JOIN batch_empleado  ON batch_empleado.ternro = cabliq.empleado "
                                sql = sql & " AND detliq.concnro = " & VectorNroACCO(G)
                            End If
                                OpenRecordset sql, rsConsult
                                If Not rsConsult.EOF Then
                                    If EsNulo(rsConsult!Monto) Then
                                        l_monto = 0
                                    Else
                                        l_monto = rsConsult!Monto
                                    End If
                                    VectorValorACCO(G) = VectorValorACCO(G) + l_monto
                                End If
                                rsConsult.Close
                        
                        End If

                    End If
                End If
           Next

            'Si Entra alguna vez graba
            GrabaEmpleado = True
            Flog.writeline "   Se encontraron datos para el empleado en el proceso "
        Else
           Flog.writeline "El empleado no se encuentra en el proceso. Nro de proceso: " & proNro
        End If
        
        ' Se guardan los datos de la tabla rep_formularioF21_det
        '---LOG---
        Flog.writeline "Buscando datos de la estructura"
        
        'Estructura AFP
        If TEAfp <> 0 And proNro <> 0 Then
            sql = "SELECT * from estructura "
            sql = sql & "WHERE tenro = " & TEAfp
            sql = sql & " ORDER BY 1"
            OpenRecordset sql, rsEstrnro
            'l_ternroanterior = 0
            While Not rsEstrnro.EOF
                l_monto_total = 0
                contempl = 0
                
                sql = "SELECT * "
                sql = sql & " FROM his_estructura"
                sql = sql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                sql = sql & " INNER JOIN batch_empleado ON batch_empleado.ternro = his_estructura.ternro "
                sql = sql & " AND his_estructura.estrnro = " & rsEstrnro!estrnro
                sql = sql & " AND htetdesde <= " & ConvFecha(fecEstr) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(fecEstr) & " OR htethasta IS NULL) "
                sql = sql & " WHERE his_estructura.tenro= " & TEAfp
                sql = sql & " AND bpronro = " & NroProceso
                OpenRecordset sql, rsConsult
                While Not rsConsult.EOF
                    sql = "SELECT almonto monto"
                    sql = sql & " FROM acu_liq"
                    sql = sql & " INNER JOIN cabliq ON cabliq.cliqnro = acu_liq.cliqnro AND empleado = " & rsConsult!ternro
                    sql = sql & " WHERE acunro = " & acumafp
                    sql = sql & " AND proNro = " & proNro
                    OpenRecordset sql, rsConsult2
                    If Not rsConsult2.EOF Then
                        If EsNulo(rsConsult2!Monto) Then
                            l_monto = 0
                        Else
                            l_monto = rsConsult2!Monto
                        End If
                        
                        'If rsConsult!ternro <> l_ternroanterior Then
                        l_monto_total = l_monto_total + l_monto
                        contempl = contempl + 1
                        'l_ternroanterior = rsConsult!ternro
                        'End If
                    End If
                    rsConsult.MoveNext
                Wend
                
                If l_monto_total > 0 Then
                    
                    'Inserto en la tabla rep_formularioF21_det
                    StrSql = " INSERT INTO rep_formularioF21_det "
                    StrSql = StrSql & " (bpronro, tenro, estrnro, tipo, valor, cantempl"
                    StrSql = StrSql & ")"
                    StrSql = StrSql & " VALUES "
                    StrSql = StrSql & "(" & NroProceso
                    StrSql = StrSql & "," & TEAfp
                    StrSql = StrSql & "," & rsEstrnro!estrnro
                    StrSql = StrSql & "," & 1
                    StrSql = StrSql & "," & numberForSQL(l_monto_total)
                    StrSql = StrSql & "," & numberForSQL(contempl)
                    StrSql = StrSql & ")"
                    
                    '------------------------------------------------------------------
                    'Guardo los datos en la BD
                    '------------------------------------------------------------------
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    Flog.writeline " Se Grabo el detalle del Formulario F21 para la estructura AFP"
                    
                End If
                rsEstrnro.MoveNext
            Wend
            rsConsult.Close
            rsEstrnro.Close
            rsConsult2.Close
        End If
        
        
        'Estructura ISAPRE
        l_monto = 0
        
        Flog.writeline " estructura Isapre" & TEIsapre
        If TEIsapre <> 0 And proNro <> 0 Then
            sql = "SELECT * from estructura "
            sql = sql & "WHERE tenro = " & TEIsapre
            sql = sql & " ORDER BY 1"
            Flog.writeline " query 1: " & sql
            OpenRecordset sql, rsEstrnro
            'l_ternroanterior = 0
            While Not rsEstrnro.EOF
                Flog.writeline " Nro de Estructuras" & rsEstrnro!estrnro
                l_monto_total = 0
                contempl = 0
                
                sql = "SELECT * "
                sql = sql & " FROM his_estructura"
                sql = sql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                sql = sql & " INNER JOIN batch_empleado ON batch_empleado.ternro = his_estructura.ternro "
                sql = sql & " AND his_estructura.estrnro = " & rsEstrnro!estrnro
                sql = sql & " AND htetdesde <= " & ConvFecha(fecEstr) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(fecEstr) & " OR htethasta IS NULL) "
                sql = sql & " WHERE his_estructura.tenro= " & TEIsapre
                sql = sql & " AND bpronro = " & NroProceso
                Flog.writeline " query 2: " & sql
                OpenRecordset sql, rsConsult
                While Not rsConsult.EOF
                    Flog.writeline " ternro " & rsConsult!ternro
                    sql = "SELECT almonto monto"
                    sql = sql & " FROM acu_liq"
                    sql = sql & " INNER JOIN cabliq ON cabliq.cliqnro = acu_liq.cliqnro AND empleado = " & rsConsult!ternro
                    sql = sql & " WHERE acunro = " & acumisapre
                    sql = sql & " AND proNro = " & proNro
                    Flog.writeline " query 3: " & sql
                    OpenRecordset sql, rsConsult2
                    If Not rsConsult2.EOF Then
                        If EsNulo(rsConsult2!Monto) Then
                            l_monto = 0
                        Else
                            l_monto = rsConsult2!Monto
                        End If
                        
                        'If rsConsult!ternro <> l_ternroanterior Then
                        l_monto_total = l_monto_total + l_monto
                        contempl = contempl + 1
                        'l_ternroanterior = rsConsult!ternro
                        'End If
                    End If
                    rsConsult.MoveNext
                Wend
                
                Flog.writeline " monto total " & l_monto_total
                If l_monto_total > 0 Then
                    
                    'Inserto en la tabla rep_formularioF21_det
                    StrSql = " INSERT INTO rep_formularioF21_det "
                    StrSql = StrSql & " (bpronro, tenro, estrnro, tipo, valor, cantempl"
                    StrSql = StrSql & ")"
                    StrSql = StrSql & " VALUES "
                    StrSql = StrSql & "(" & NroProceso
                    StrSql = StrSql & "," & TEIsapre
                    StrSql = StrSql & "," & rsEstrnro!estrnro
                    StrSql = StrSql & "," & 2
                    StrSql = StrSql & "," & numberForSQL(l_monto_total)
                    StrSql = StrSql & "," & numberForSQL(contempl)
                    StrSql = StrSql & ")"
                    Flog.writeline " query 4: " & StrSql
                    '------------------------------------------------------------------
                    'Guardo los datos en la BD
                    '------------------------------------------------------------------
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    Flog.writeline " Se Grabo el detalle del Formulario F21 para la estructura Isapre"
                    
                End If
                rsEstrnro.MoveNext
            Wend
            rsConsult.Close
            rsEstrnro.Close
            rsConsult2.Close
        End If
Next
        
  
If GrabaEmpleado Then
  
    '------------------------------------------------------------------
    'Armo la SQL para guardar los datos
    '------------------------------------------------------------------

    StrSql = " INSERT INTO rep_formularioF21 "
    StrSql = StrSql & " (bpronro, pronro, prodesc, descripcion, fecha, hora, iduser,"
    StrSql = StrSql & " empnro, empnom, empdir, emprut, pliqnro, pliqmes, pliqanio"
    For I = 1 To 25
        StrSql = StrSql & "," & " valor" & I
    Next I
    
    StrSql = StrSql & ")"
    StrSql = StrSql & " VALUES "
    StrSql = StrSql & "(" & NroProceso
    StrSql = StrSql & "," & proNro
    StrSql = StrSql & ",'" & prodesc & "'"
    StrSql = StrSql & ",'" & Mid(descripcion, 1, 98) & "'"
    StrSql = StrSql & ",'" & Fecha & "'"
    StrSql = StrSql & ",'" & Hora & "'"
    StrSql = StrSql & ",'" & IdUser & "'"
    StrSql = StrSql & "," & Empnro & ""
    StrSql = StrSql & ",'" & empresa & "'"
    StrSql = StrSql & ",'" & emprDire & "'"
    StrSql = StrSql & ",'" & emprCuit & "'"
    StrSql = StrSql & "," & pliqNro
    StrSql = StrSql & "," & pliqMes
    StrSql = StrSql & "," & pliqAnio
    For I = 1 To 25
        StrSql = StrSql & "," & numberForSQL(VectorValorACCO(I))
    Next I
    StrSql = StrSql & ")"
    
    '------------------------------------------------------------------
    'Guardo los datos en la BD
    '------------------------------------------------------------------
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline " Se Grabo el Formulario F21"
End If

Exit Sub

MError:
    Flog.writeline "Error en Formulario F21: " & NroProceso & " Error: " & Err.Description
    Flog.writeline "Ultimo Sql Ejecutado: " & StrSql
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub



'--------------------------------------------------------------------
' Se encarga de generar un ResultSet de los empleados a cambiar
' si el RS es vacio significa que hay que aplicarlo sobre todos
'--------------------------------------------------------------------
Sub CargarEmpleados(ByVal NroProc, ByRef rsEmpl As ADODB.Recordset, ByVal empresa As Long)

Dim StrEmpl As String

    If NroProc > 0 Then
        StrEmpl = "SELECT * FROM batch_empleado "
        StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = batch_empleado.ternro "
        StrEmpl = StrEmpl & " WHERE bpronro = " & NroProc
        StrEmpl = StrEmpl & " ORDER BY progreso,estado"
    End If
   
    OpenRecordset StrEmpl, rsEmpl
    
    cantRegistros = rsEmpl.RecordCount
    totalEmpleados = cantRegistros
    
End Sub

Function numberForSQL(Str)
   
  numberForSQL = Replace(Str, ",", ".")

End Function


Function strForSQL(Str)
   
  If IsNull(Str) Then
     strForSQL = "NULL"
  Else
     strForSQL = Str
  End If

End Function



Public Function Calcular_Edad(ByVal Fecha As Date, ByVal Hasta As Date) As Integer
'...........................................................................
' Archivo       : edad.i                              fecha ini. : 20/01/92
' Nombre progr. :
' tipo programa : FGZ
' Descripcion   :
'...........................................................................
Dim años  As Integer
Dim ALaFecha As Date

    ALaFecha = C_Date(Hasta)
    
    años = Year(ALaFecha) - Year(Fecha)
    If Month(ALaFecha) < Month(Fecha) Then
       años = años - 1
    Else
        If Month(ALaFecha) = Month(Fecha) Then
            If Day(ALaFecha) < Day(Fecha) Then
                años = años - 1
            End If
        End If
    End If
    Calcular_Edad = años
End Function


Sub buscarDatosEmpresa(Empnroestr)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset

    empresa = ""
    emprTer = 0
    emprCuit = ""
    emprDire = ""
    
    ' -------------------------------------------------------------------------
    'Busco los datos Basicos de la Empresa
    ' -------------------------------------------------------------------------
    Flog.writeline "Buscando datos de la empresa"
    
    StrSql = "SELECT * FROM empresa WHERE Estrnro = " & Empnroestr
    OpenRecordset StrSql, rsConsult
    
    If rsConsult.EOF Then
       Flog.writeline "Error: Buscando datos de la empresa: al obtener el empleado"
       HuboErrores = True
    Else
        empresa = rsConsult!empnom
        emprTer = rsConsult!ternro
        Empnro = rsConsult!Empnro
    End If
    
    rsConsult.Close
            
    'Consulta para obtener el RUT de la empresa
    StrSql = "SELECT nrodoc FROM tercero " & _
             " INNER JOIN ter_doc ON (tercero.ternro = ter_doc .ternro and ter_doc.tidnro = 1)" & _
             " Where tercero.ternro =" & emprTer
    
    Flog.writeline "Buscando datos del RUT de la empresa"
    
    OpenRecordset StrSql, rsConsult
    
    If rsConsult.EOF Then
        Flog.writeline "No se encontró el RUT de la Empresa"
        emprCuit = "  "
    Else
        emprCuit = rsConsult!nrodoc
    End If
    rsConsult.Close
End Sub

