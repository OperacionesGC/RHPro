Attribute VB_Name = "repLibroRemu"
'Global Const Version = "1.00" ' Diego Rosso
'Global Const FechaModificacion = "19/11/2007"
'Global Const UltimaModificacion = "" 'Version Inicial

'Global Const Version = "1.01" ' Martin Ferraro
'Global Const FechaModificacion = "31/07/2009"
'Global Const UltimaModificacion = "" 'Encriptacion de string connection

'Global Const Version = "1.02" ' Carmen Quintero
'Global Const FechaModificacion = "30/11/2011"
'Global Const UltimaModificacion = "" 'Se modificó para que en la columna 41 se pueda colocar CO ó AC

Global Const Version = "1.03" ' Miriam Ruiz - CAS-26972 - H&A - Error en el Mapeo de Documentos en el libro de remuneraciones
Global Const FechaModificacion = "28/10/2014"
Global Const UltimaModificacion = "" 'se agregó el join con la tabla tipodocu_pais

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
Global emprTer As Long
Global emprDire As String
Global emprCuit

Global listapronro       'Lista de procesos

Global totalEmpleados
Global cantRegistros

Global incluyeAgencia As Integer
Global NroAcDiasTrabajados As Long

Global TipoEstructura As Long

Global VectorEsConc(40) As Boolean 'True = Concepto    False = Acumulador
Global VectorNroACCO(40) As Long
Global VectorValorACCO(40) As Double

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

Dim IdUser As String
Dim Fecha As Date
Dim Hora As String

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
    
    TiempoInicialProceso = GetTickCount
    
    Nombre_Arch = PathFLog & "ReporteRemuneraciones" & "-" & NroProceso & ".log"
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

       'Obtengo la lista de procesos
       listapronro = ArrParametros(0)

       'Obtengo el modelo a usar para obtener los datos
       tipoModelo = ArrParametros(1)


       'Obtengo el nro de la ultima pagina impresa
       Pagina = CLng(ArrParametros(2))

       
       'Obtengo el numero de empresa
       If CLng(ArrParametros(3)) <> 0 Then
            Empnroestr = CLng(ArrParametros(3))
       Else
            Flog.writeline "No Se selecciono el parametro Empresa. "
            HuboErrores = True
       End If
       
       'Obtengo los cortes de estructura
       tenro1 = CInt(ArrParametros(4))
       estrnro1 = CInt(ArrParametros(5))

       tenro2 = CInt(ArrParametros(6))
       estrnro2 = CInt(ArrParametros(7))

       tenro3 = CInt(ArrParametros(8))
       estrnro3 = CInt(ArrParametros(9))

       If UBound(ArrParametros) > 9 Then
        fecEstr = ArrParametros(10)
       End If

       'Obtengo el titulo del reporte
       If UBound(ArrParametros) > 10 Then
        tituloReporte = ArrParametros(11)
       Else
        tituloReporte = ""
       End If

      
       'Obtengo el numero de pliqnro
       
        If ArrParametros(12) <> 0 Then
            pliqNro = CLng(ArrParametros(12))
        Else
            Flog.writeline "No se selecciono un Periodo."
            HuboErrores = True
        End If
       
             
        
        'EMPIEZA EL PROCESO
        Flog.writeline "Generando el reporte"
                
       
        
        'INICIALIZ0 ARRAY DE CONCEPTOS
        '************************
        For I = 1 To 40
            VectorEsConc(I) = False 'As Boolean
            VectorNroACCO(I) = 0 ' As Long
            VectorValorACCO(I) = 0 'As Double
        Next
            
        NroAcDiasTrabajados = 0
        TipoEstructura = 0
     
        'Obtengo la Configuracion del Confrep
        StrSql = "SELECT * FROM confrep"
        StrSql = StrSql & " WHERE repnro = 219"
        OpenRecordset StrSql, objRs2
        
        Flog.writeline "Obtengo los datos del confrep"
        
        If objRs2.EOF Then
          Flog.writeline "No esta configurado el ConfRep para el reporte"
          HuboErrores = True
        End If
       
       
        
       If Not HuboErrores Then
       
            Do Until objRs2.EOF
                 ' Levanto parametros del Confrep     SON 40 CONCEPTOS/ACUMULADORES
               Select Case objRs2!confnrocol
                  Case 1 To 40
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
                     Else
                        VectorEsConc(objRs2!confnrocol) = False
                        VectorNroACCO(objRs2!confnrocol) = objRs2!confval
                     End If
                  Case 41
                     If objRs2!conftipo = "AC" Then
                         NroAcDiasTrabajados = objRs2!confval
                     Else
                        If objRs2!conftipo = "CO" Then
                            StrSql = "SELECT concnro FROM concepto WHERE conccod = " & objRs2!confval
                            StrSql = StrSql & " OR conccod = '" & objRs2!confval2 & "'"
                            OpenRecordset StrSql, objRs3
                            If Not objRs3.EOF Then
                                NroAcDiasTrabajados = CLng(objRs3!concnro)
                            End If
                            objRs3.Close
                        Else
                            Flog.writeline "La columna 41 Dias trabajados debe ser un acumulador o un concepto. "
                        End If
                     End If
                  Case 42
                     If objRs2!conftipo = "TE" Then
                         TipoEstructura = objRs2!confval
                     Else
                         Flog.writeline "La columna 42 Dias trabajados debe ser un Tipo de Estructura. "
                     End If
                
                  Case Else
                       
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
                    
       
              'Grabo Cabecera ***************************************************************
              '******************************************************************************
                StrSql = " INSERT INTO rep_libroremu "
                StrSql = StrSql & " (bpronro , fecha,hora, iduser, empnro, empnom, empdir,"
                StrSql = StrSql & " emprut , pliqnro, pliqmes, pliqanio,"
                StrSql = StrSql & " cantemp , ultima_pag_impr) "
                StrSql = StrSql & " VALUES "
                StrSql = StrSql & "(" & NroProceso
                StrSql = StrSql & "," & ConvFecha(Fecha)
                StrSql = StrSql & ",'" & Format(Hora, "hh:mm:ss")
                StrSql = StrSql & "','" & IdUser
                StrSql = StrSql & "'," & Empnro
                StrSql = StrSql & ",'" & empresa
                StrSql = StrSql & "','" & emprDire
                StrSql = StrSql & "','" & emprCuit
                StrSql = StrSql & "'," & pliqNro
                StrSql = StrSql & "," & pliqMes
                StrSql = StrSql & "," & pliqAnio
                StrSql = StrSql & "," & cantRegistros
                StrSql = StrSql & "," & Pagina
                StrSql = StrSql & ")"
    
                '------------------------------------------------------------------
                'Guardo los datos en la BD
                '------------------------------------------------------------------
                 objConn.Execute StrSql, , adExecuteNoRecords
       
       
             'Obtengo la lista de procesos
             arrpronro = Split(listapronro, ",")

             ord = 0
     
            'Genero por cada empleado un registro
            Do Until rsEmpl.EOF
                arrpronro = Split(listapronro, ",")
                EmpErrores = False
                ternro = rsEmpl!ternro
                orden = ord
                 Flog.writeline ""
                 Flog.writeline "Generando datos empleado " & ternro
                        
                 Call ReporteRemuneracion(arrpronro, ternro, tituloReporte, orden)
                          
                                            
                'Actualizo el estado del proceso
                TiempoAcumulado = GetTickCount
                
                'Resto uno a la cantidad de registros
                cantRegistros = cantRegistros - 1
                
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
                
                'Proximo Empleado
                rsEmpl.MoveNext
            Loop
           
                'Si se llega dar el caso de que no se grabo ningun empleado se borra la cabecera.
                  If CantEmpGrabados = 0 Then
                     StrSql = "DELETE  FROM rep_libroremu "
                     StrSql = StrSql & " WHERE bpronro = " & NroProceso
                     objConn.Execute StrSql, , adExecuteNoRecords
                  End If
                        
            End If
    Else
        
        Exit Sub
    End If
   
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline
    Flog.writeline "************************************************************"
    Flog.writeline "Fin :" & Now
    Flog.writeline "Cantidad de empleados guardados en el reporte: " & CantEmpGrabados
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
Sub ReporteRemuneracion(ListaPro, ternro As Long, descripcion As String, orden As Long)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset

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


On Error GoTo MError

Flog.writeline "Busco los datos del empleado "

'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleg,terape,terape2,ternom,ternom2 "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & ternro

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   nombre = rsConsult!ternom
   If IsNull(rsConsult!ternom2) Then
      nombre2 = ""
   Else
      nombre2 = rsConsult!ternom2
   End If
   apellido = rsConsult!terape
   If IsNull(rsConsult!terape2) Then
      apellido2 = ""
   Else
      apellido2 = rsConsult!terape2
   End If
   Legajo = rsConsult!empleg
Else
   Flog.writeline "Error al obtener los datos del empleado"
   GoTo MError
End If

rsConsult.Close



'------------------------------------------------------------------
' Obtengo los datos de la estructura configurada en el confrep
'------------------------------------------------------------------

'---LOG---
Flog.writeline "Buscando datos de la estructura"

If TipoEstructura <> 0 Then
   sql = " SELECT estrdabr "
   sql = sql & " FROM estructura "
   sql = sql & " INNER JOIN his_estructura ON his_estructura.estrnro = estructura.estrnro AND his_estructura.tenro = " & TipoEstructura
   sql = sql & " AND his_estructura.htetdesde <= " & ConvFecha(fecEstr) & " AND (his_estructura.htethasta IS NULL OR his_estructura.htethasta >= " & ConvFecha(fecEstr) & ")"
   sql = sql & " WHERE his_estructura.ternro = " & ternro
   OpenRecordset sql, rsConsult
   
   If Not rsConsult.EOF Then
        If IsNull(rsConsult!estrdabr) Then
           DescEstructura = ""
        Else
          DescEstructura = rsConsult!estrdabr
          Flog.writeline "Estructura Obtenida"
        End If
    Else
       DescEstructura = ""
   End If
   rsConsult.Close
Else
   DescEstructura = ""
End If




'*************************************************************************
'Consulta para obtener el RUT del empleado
'*************************************************************************
Flog.writeline "Buscando datos del RUT de la empresa"


StrSql = "SELECT nrodoc FROM tercero " & _
         " INNER JOIN ter_doc ON tercero.ternro = ter_doc .ternro" & _
         " inner join tipodocu_pais ON ter_doc.tidnro=tipodocu_pais.tidnro and tipodocu_pais.tidcod=1 and tipodocu_pais.paisnro=8 " & _
         " Where tercero.ternro = " & ternro
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
    Flog.writeline "No se encontró el RUT del Empleado"
    RUT = ""
Else
    RUT = rsConsult!nrodoc
End If
rsConsult.Close



'Inicializo Conceptos
For I = 1 To 40
   VectorValorACCO(I) = 0 'As Double
Next

DiasTrabajados = 0

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
         '  prodesc = rsConsult!prodesc
         '  pliqdesc = rsConsult!pliqdesc
           cliqnro = rsConsult!cliqnro
           pliqdesde = rsConsult!pliqdesde
           pliqhasta = rsConsult!pliqhasta
        
            rsConsult.Close
        
            '------------------------------------------------------------------
            'Busco los datos de los 40 acumuladores/conceptos
            '------------------------------------------------------------------
            'VectorEsConc(40) As Boolean 'True = Concepto    False = Acumulador
            'VectorNroACCO(40) As Long  Numero de acumulador o concepto
            'VectorValorACCO(40) As Double
            
            
            For G = 1 To 40
                If VectorNroACCO(G) <> 0 Then 'Si es igual a cero es porq no se configuro valor para ese AC/CO
                    If VectorEsConc(G) = False Then
                        sql = " SELECT almonto "
                        sql = sql & " FROM acu_liq"
                        sql = sql & " WHERE acunro = " & VectorNroACCO(G)
                        sql = sql & " AND cliqnro =  " & cliqnro
                    Else
                        sql = " SELECT detliq.dlimonto AS almonto "
                        sql = sql & " FROM cabliq "
                        sql = sql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND cabliq.pronro = " & proNro
                        sql = sql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro "
                        sql = sql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & ternro & " AND detliq.concnro = " & VectorNroACCO(G)
                    End If
                    OpenRecordset sql, rsConsult
                    
                    If Not rsConsult.EOF Then
                        VectorValorACCO(G) = VectorValorACCO(G) + rsConsult!almonto
                    End If
                    rsConsult.Close
                End If
            Next
            
                 
            'Buscar valor dias trabajados
            If NroAcDiasTrabajados <> 0 Then
                sql = " SELECT almonto "
                sql = sql & " FROM acu_liq"
                sql = sql & " WHERE acunro = " & NroAcDiasTrabajados
                sql = sql & " AND cliqnro =  " & cliqnro
                OpenRecordset sql, rsConsult
                If Not rsConsult.EOF Then
                    DiasTrabajados = rsConsult!almonto
                End If
                
                If DiasTrabajados = 0 Then
                    rsConsult.Close
                    sql = " SELECT detliq.dlimonto AS almonto "
                    sql = sql & " FROM cabliq "
                    sql = sql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND cabliq.pronro = " & proNro
                    sql = sql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro "
                    sql = sql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & ternro & " AND detliq.concnro = " & NroAcDiasTrabajados
                    OpenRecordset sql, rsConsult
                    If Not rsConsult.EOF Then
                        DiasTrabajados = rsConsult!almonto
                    End If
                End If
                rsConsult.Close
            Else
                DiasTrabajados = 0
            End If
             'Si Entra alguna vez graba
             GrabaEmpleado = True
             Flog.writeline "   Se encontraron datos para el empleado en el proceso "
        Else
           Flog.writeline "El empleado no se encuentra en el proceso. Nro de proceso: " & proNro
        End If
Next
        
  
If GrabaEmpleado Then
  
    '------------------------------------------------------------------
    'Armo la SQL para guardar los datos
    '------------------------------------------------------------------
    
    
    StrSql = " INSERT INTO rep_libroremu_det "
    StrSql = StrSql & " (bpronro, "
    StrSql = StrSql & " ternro , empleg, terape, terape2,"
    StrSql = StrSql & " ternom , ternom2, rut, diastrab,"
    StrSql = StrSql & " puesto"
    For I = 1 To 40
        StrSql = StrSql & "," & " valor" & I
    Next I
    StrSql = StrSql & ")"
    StrSql = StrSql & " VALUES "
    StrSql = StrSql & "(" & NroProceso
    StrSql = StrSql & "," & ternro
    StrSql = StrSql & "," & Legajo
    StrSql = StrSql & ",'" & apellido & "'"
    StrSql = StrSql & ",'" & apellido2 & "'"
    StrSql = StrSql & ",'" & nombre & "'"
    StrSql = StrSql & ",'" & nombre2 & "'"
    StrSql = StrSql & ",'" & RUT & "'"
    StrSql = StrSql & "," & DiasTrabajados
    StrSql = StrSql & ",'" & DescEstructura & "'"
    For I = 1 To 40
        StrSql = StrSql & "," & VectorValorACCO(I)
    Next I
    StrSql = StrSql & ")"
    
    '------------------------------------------------------------------
    'Guardo los datos en la BD
    '------------------------------------------------------------------
    objConn.Execute StrSql, , adExecuteNoRecords
    
     Flog.writeline "   Se Grabo el empleado."
    'Sumo uno a la cantidad de empleados guardados
    CantEmpGrabados = CantEmpGrabados + 1
End If

Exit Sub

MError:
    Flog.writeline "Error en empleado: " & Legajo & " Error: " & Err.Description
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
        'If incluyeAgencia = -1 Then
           'StrEmpl = StrEmpl & " AND (agencia is null OR agencia = 0)"
        'End If
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
        
    'Consulta para obtener la direccion de la empresa
    StrSql = "SELECT detdom.calle,detdom.nro,localidad.locdesc, detdom.piso, detdom.oficdepto From cabdom " & _
        " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & emprTer & _
        " INNER JOIN localidad ON detdom.locnro = localidad.locnro "

    Flog.writeline "Buscando datos de la direccion de la empresa"
    
    OpenRecordset StrSql, rsConsult
    
    If rsConsult.EOF Then
        Flog.writeline "No se encontró el domicilio de la empresa"
        'Exit Sub
        emprDire = "   "
    Else
        emprDire = rsConsult!calle & " " & rsConsult!Nro
        If Not EsNulo(rsConsult!piso) Then
            emprDire = emprDire & " P. " & rsConsult!piso
        End If
        If Not EsNulo(rsConsult!oficdepto) Then
            emprDire = emprDire & " Dpto. " & rsConsult!oficdepto
        End If
        emprDire = emprDire & " - " & rsConsult!locdesc
    End If
    rsConsult.Close
    
    'Consulta para obtener el RUT de la empresa
    
    StrSql = "SELECT nrodoc FROM tercero " & _
         " INNER JOIN ter_doc ON tercero.ternro = ter_doc .ternro" & _
         " inner join tipodocu_pais ON ter_doc.tidnro=tipodocu_pais.tidnro and tipodocu_pais.tidcod=1 and tipodocu_pais.paisnro=8 " & _
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

