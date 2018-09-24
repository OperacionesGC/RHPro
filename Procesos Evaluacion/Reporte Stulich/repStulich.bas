Attribute VB_Name = "repStulich"
' ---------------------------------------------------------------------------------------------
' Descripcion: Reporte Stulich
' Autor      : Leticia Amadio.
' Fecha      : --/05/2005
' Ultima Mod.:
' Descripcion:
' Modificacion:   -05 - 2006 - LA. - tener en cuenta que se entre fecha de antiguedad en la categoria
'               18-05 - 2006 - LA. sacar sueldo de ammonto y ver si esta fuera de convenio, con el string 'Fuera de Convenio'
'               31-05-2006 - LA. Tener en cuenta para la escala de stulich una 4ª coordenada (graduado - no graduado)
' ---------------------------------------------------------------------------------------------

'Global Const Version = "1.00"
'Global Const FechaModificacion = " -05-2005 " ' Leticia Amadio
'Global Const UltimaModificacion = " " 'reporte Stulich


'Global Const Version = "1.01"
'Global Const FechaModificacion = "25-06-2007 " ' Leticia Amadio
'Global Const UltimaModificacion = " " 'incluir distintas opciones para la calificacion final ( objs Indiv o Objs gral o Calif rdp para el calculo de la calific preliminar)
                                      ' NO se tiene en cuenta la Calific Ajustada, en su lugar se usa la Calificacion Preliminar (la que se saco en la evaluacion)
    
'Global Const Version = "1.02"
'Global Const FechaModificacion = "05-07-2007 " ' Leticia Amadio
'Global Const UltimaModificacion = " " ' sacar vistas (v_), Reemplazar coma por punto en los totales, Sacar las ctes cableadas de Competencs BBca
    
'Global Const Version = "1.03"
'Global Const FechaModificacion = "14-11-2007 " ' Leticia Amadio
'Global Const UltimaModificacion = " " ' Buscar las Evaluacines de las Areas tanto si la seccion es de Areas, como Areas y Competencias del Consejero.
    
    
'Global Const Version = "1.04"
'Global Const FechaModificacion = "14-05-2008 " ' Leticia Amadio
'Global Const UltimaModificacion = " " ' Cambiar forma de mostrar resultados, se muestran comp tecnicas, compartidas, objs, calificac gral en escala de  1-100, se incluyo porcentajes a Comp Tec, Compt Comp, Objs/Calif Gral, se busca informacion de la nueva seccion Ev. Gral.
    
Global Const Version = "1.05"
Global Const FechaModificacion = "20/08/2009" ' Martin Ferraro
Global Const UltimaModificacion = " " ' Encriptacion de archivos
    
   
' ___________________________________________________________________________________________________



Option Explicit
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

Global tipoHora As String
Global TipoDia As Integer
Global FechaDesde As Date
Global FechaHasta As Date

Global Tenro1 As Long
Global Tenro2 As Long
Global Tenro3 As Long
Global Estrnro1 As Long
Global Estrnro2 As Long
Global Estrnro3 As Long
Global EmpEstrnro1 As Long
Global EmpEstrnro2 As Long
Global EmpEstrnro3 As Long
Global Categ_Estrnro As Long
Global Categ_fecha As Date
Global NroGrilla As Long


Global Evento
Global tipoEv
Global porcCComp
Global porcCTec
Global porcGral

Global Consejero
Global Aconsejado
Global TodosGrupo
Global TodosDepto
Global TodosCateg
Global TodosConsej
Global TodosAconsej

    ' para sacar el sueldo
Global Acu
Global CO
    ' xx
Global cevaluador
Global cconsejero
 ' para la calificac Ajustada ---
' Dim cantEmpls(5)

Global rsConsult As New ADODB.Recordset
Global rsConsult2 As New ADODB.Recordset
Global rsConsult3 As New ADODB.Recordset
Global rsConsult4 As New ADODB.Recordset
Global rsConsult5 As New ADODB.Recordset
Global rsConsult6 As New ADODB.Recordset

'DATOS DE LA TABLA batch_proceso
Global bpfecha As Date
Global bphora As String
Global bpusuario As String

Global repNro As Long
Global idUser As String
Global Tercero As Long
Global USA_DEBUG As Boolean


Global Puntaje As Integer       'Calif. Final Ajustada
Global Posicion As String       'Antig. "," Puntaje
Global Remuneracion As Double   'Remuneracion Maxima

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
Dim tipoDepuracion
'Dim historico As Boolean
Dim param
Dim Ternro
Dim rsEmpl As New ADODB.Recordset
Dim i
Dim totalEmpleados
Dim cantRegistros
Dim PID As String
Dim ArrParametros
Dim Parametros As String


'Dim CabAprobada As String

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

    Nombre_Arch = PathFLog & "ReporteStulich" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    HuboErrores = False
    USA_DEBUG = True
    
    
    'Obtengo la cantidad de empledos a procesar
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    cantRegistros = CInt(objRs!bprcempleados)
    totalEmpleados = cantRegistros
    
    objRs.Close
   
    Flog.writeline "Inicio Proceso del Reporte Stulich: " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    
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
       Parametros = objRs!bprcparam
       ArrParametros = Split(Parametros, ".")
       
       'Obtengo el Tipo de Hora     'TipoDia = ArrParametros(0)
       'Obtengo el Tipo de Dia     'tipoHora = ArrParametros(1)
       
        ' Obtengo el nro de reporte
       repNro = CInt(ArrParametros(0))
       
       ' Datos Filtro
       Evento = CInt(ArrParametros(1)) ' Obtengo el evento RDP
       Consejero = CInt(ArrParametros(2))
       Aconsejado = CInt(ArrParametros(3))
        'Obtengo las estructuras
       Tenro1 = CInt(ArrParametros(4))
       Estrnro1 = CInt(ArrParametros(5))
       Tenro2 = CInt(ArrParametros(6))
       Estrnro2 = CInt(ArrParametros(7))
       Tenro3 = CInt(ArrParametros(8))
       Estrnro3 = CInt(ArrParametros(9))
       Categ_fecha = ArrParametros(10)
       tipoEv = CInt(ArrParametros(11)) ' tipo de Evaluacion a tomar para los calculos - (Objs Indiv - Obj Gral - RDP)
       porcCComp = CInt(ArrParametros(12))
       porcCTec = CInt(ArrParametros(13))
       porcGral = CInt(ArrParametros(14))
       
              
              
       'Obtengo las fechas
       'FechaDesde = objRs!bprcfecdesde  'FechaHasta = objRs!bprcfechasta
       
       'Obtengo la fecha del proceso
       bpfecha = objRs!bprcfecha
       'Obtengo la hora del proceso
       bphora = objRs!bprchora
       
       'Obtengo el usuario del proceso
       bpusuario = objRs!idUser
       
        
        Flog.writeline " Parametros que entraron: "
        Flog.writeline "    Evento: " & Evento
        Flog.writeline "    Consejero: " & Consejero
        Flog.writeline "    Aconsejado: " & Aconsejado
        Flog.writeline "    Tenro1: " & Tenro1
        Flog.writeline "    Estrnro1: " & Estrnro1
        Flog.writeline "    Tenro2: " & Tenro2
        Flog.writeline "    Estrnro2: " & Estrnro2
        Flog.writeline "    Tenro3: " & Tenro3
        Flog.writeline "    Estrnro3: " & Estrnro3
        Flog.writeline "    Fecha de la Categoría: " & Categ_fecha
        Flog.writeline "    Tipo de Resultado Final: " & tipoEv
        Flog.writeline "    Porcentaje para Competencias Comp.: " & porcCComp
        Flog.writeline "    Porcentaje para Competencias Téc.: " & porcCTec
        Flog.writeline "    Porcentaje para Calificacion Gral/ Objs : " & porcGral
        
        
        
       'EMPIEZA EL PROCESO
        StrSql = " DELETE  rep_stulich "
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
       
       'Cargo los datos del confrep - De aca sale el de dde se saca el sueldo (concepto - ..) y cevaluador
        Call cargarConfRep(repNro, Acu, CO, cevaluador, cconsejero)
        
        ' se fija si se seleccionaron opciones del filtro.
       Call datosFiltro(Aconsejado, Consejero, Estrnro1, Estrnro2, Estrnro3)
       
       'Obtengo los empleados sobre los que tengo que generar la evaluac Final.
       Call CargarEmpleados(NroProceso, rsEmpl)
       
       StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
       objConn.Execute StrSql, , adExecuteNoRecords
       
       'Genero por cada empleado -------
       Do Until rsEmpl.EOF
         'MyBeginTrans
         Ternro = rsEmpl!Ternro   ' Ternro = Aconsejado
         
         Flog.writeline " "
         Flog.writeline "Se generan los datos del reporte Stulich para el Tercero " & Ternro
         Call generarDatos(Ternro) ', FechaDesde, FechaHasta, TipoDia, tipoHora
         
          'Actualizo el estado del proceso
         TiempoAcumulado = GetTickCount
           
         cantRegistros = cantRegistros - 1
        
        If totalEmpleados <> 0 Then
         StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                    ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                    ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
          objConn.Execute StrSql, , adExecuteNoRecords
         End If
          'Si se generaron todos los recibos de sueldo del empleado correctamente lo borro
'          If Not EmpErrores Then
'              StrSql = " DELETE FROM batch_empleado "
'              StrSql = StrSql & " WHERE bpronro = " & NroProceso
'              StrSql = StrSql & " AND ternro = " & Ternro
'              objConn.Execute StrSql, , adExecuteNoRecords
'          End If
          'MyCommitTrans
          rsEmpl.MoveNext
       Loop
                
       ' NO se usa mas la Calific. Ajustada
       ' Call DatosCalifAjustada(NroProceso)
       
       ' FGZ - 18/08/2005
       Call AjustarPosicion_Y_Remuneracion
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
    
    objConn.Close
    Flog.writeline "Fin :" & Now
    Flog.Close
    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now
End Sub


Function controlNull(Str)
  If Trim(Str) = "" Then
     controlNull = "null"
  Else
     controlNull = "'" & Str & "'"
  End If
End Function

'--------------------------------------------------------------------
' Se encarga de generar un ResultSet de los empleados a cambiar
' si el RS es vacio significa que hay que aplicarlo sobre todos
'--------------------------------------------------------------------
Sub CargarEmpleados(NroProc, ByRef rsEmpl As ADODB.Recordset)

Dim StrEmpl As String

    StrEmpl = "SELECT * FROM batch_empleado WHERE bpronro = " & NroProc & " ORDER BY progreso"
    
    OpenRecordset StrEmpl, rsEmpl
End Sub


'FUNCION: Convierte un string que contiene una hora al formato string hora
Function convHora(Str)
  convHora = Mid(Str, 1, 2) & ":" & Mid(Str, 3, 2)
End Function 'convHora(str)

' ________________________________________________________________
' ________________________________________________________________
Sub cargarConfRep(nroRep, ByRef Acu, ByRef CO, ByRef cevaluador, ByRef cconsejero)
Dim objRs As New ADODB.Recordset
Dim EncontroGrilla As Boolean

On Error GoTo ME_ConfRep

    StrSql = " SELECT  conftipo,  confetiq, confval, confval2 "  'confnrocol,confaccion,
    StrSql = StrSql & " FROM confrep "
    StrSql = StrSql & " WHERE repnro = " & nroRep
    'StrSql = StrSql & " AND conftipo = 'TH'"

    OpenRecordset StrSql, objRs
    
    'inicializar
    EncontroGrilla = False
    Do Until objRs.EOF
        If UCase(objRs!conftipo) = "ACU" Then
            Acu = objRs!confval
        Else
            If UCase(objRs!conftipo) = "CO" Then
                CO = objRs!confval
            Else
                If UCase(objRs!conftipo) = "GRI" Then
                    NroGrilla = objRs!confval
                    EncontroGrilla = True
                Else
                    If UCase(objRs!confval2) = "CONSEJERO" Then
                    cconsejero = objRs!confval
                    Else
                    cevaluador = objRs!confval
                    End If
                End If
            End If
        End If

     '   Columnas(CLng(objRs!confnrocol)) = agregar(Columnas(CLng(objRs!confnrocol)), objRs!confval & "@" & objRs!confaccion)
     '   ColumnasTitulos(CLng(objRs!confnrocol)) = objRs!confetiq
        
      objRs.MoveNext
    Loop
    
    If Not EncontroGrilla Then
        Flog.writeline "El nro de grilla no esta configurada"
        NroGrilla = 0
    End If
    
    
    objRs.Close
    
Exit Sub

ME_ConfRep:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub 'cargarConfRep

' _________________________________________________________________
' setea las variables para saber que filtro se eligio ....
' _________________________________________________________________
Sub datosFiltro(Aconsejado, Consejero, Estrnro1, Estrnro2, Estrnro3)

If Aconsejado = 0 Then
    TodosAconsej = -1
Else
    TodosAconsej = Aconsejado
End If

If Consejero = 0 Then
    TodosConsej = -1
Else
    TodosConsej = Consejero
End If
If Estrnro1 = 0 Then
    TodosGrupo = -1
Else
    TodosGrupo = Estrnro1
End If
If Estrnro2 = 0 Then
    TodosDepto = -1
Else
    TodosDepto = Estrnro2
End If
If Estrnro3 = 0 Then
    TodosCateg = -1
Else
    TodosCateg = Estrnro3
End If

End Sub

'--------------------------------------------------------------------
' Se encarga de buscar todos los datos referente a un empleado
'--------------------------------------------------------------------
Sub generarDatos(ByVal Ternro As Integer) ' , ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByVal TipoDia As Integer, TipoHoras As String

Dim StrSql As String
Dim Campos As String
Dim Valores As String

Dim rsConsult1 As New ADODB.Recordset
Dim rsConsult2 As New ADODB.Recordset
Dim Legajo As Integer
Dim Cantidad As Integer
Dim i

Dim empleg
Dim ApeyNom As String
Dim FecNac As Date
Dim Edad As Integer
Dim AntigCat As String
Dim AntigNum As Integer
Dim sueldo
Dim ConsejeroDesc As String
Dim Convenio As String
Dim EmpAnt As String

Dim Estrnro1Desc
Dim Estrnro2Desc
Dim Estrnro3Desc

Dim AreaLetra(100)
Dim Area(100)
Dim AreaVal(100)
Dim AreaMax(100)
Dim CantAreas
Dim CalifRDP
Dim CalifGralRDP
Dim CalifGralRDPPorc
Dim ObjsCant(100)
Dim objsVal(100)
'Dim ObjsResu(100)
Dim cantObjs
Dim totalObjGral
Dim cantObjGral ' representa la cantidad de RDEs con Objs Generales evaluados
Dim compCentral
Dim compCompart
Dim compTecPorc
Dim compCompPorc
Dim ObjsPorc
Dim ObjGralPorc
Dim totalObjs
Dim totalComp
Dim totalPuntos
Dim evalFinal
Dim califPrel
Dim tipoEvGral
Dim resuEvaGral

On Error GoTo MError

' ++++++++++++++++++++++++++
    ' FechaActual = CDate(FechaDesde) ' Contador = DateDiff("d", FechaDesde, FechaHasta)
    'Busco los datos del tipos de estructura 1, 2 y 3

Call DatosEstructura(Ternro, EmpEstrnro1, EmpEstrnro2, EmpEstrnro3, Estrnro1Desc, Estrnro2Desc, Estrnro3Desc)
Call DatosEmp(Ternro, empleg, ApeyNom, FecNac, Edad)
Call bus_AntEstructura(Ternro, EmpEstrnro3, Categ_fecha, AntigCat, AntigNum)
Call DatosVarios(Ternro, Convenio, EmpAnt)
Call DatosConsejero(Ternro, Consejero, ConsejeroDesc)
Call DatosSueldo(Ternro, Acu, CO, sueldo)
Call DatosArea(Ternro, Area, AreaLetra, AreaVal, CantAreas, AreaMax)
Call DatosRDP(Ternro, CalifRDP, CalifGralRDP, CalifGralRDPPorc)
Call DatosObj(Ternro, ObjsCant, objsVal, cantObjs)
Call DatosObjGral(Ternro, totalObjGral, cantObjGral)
Call DatosResuArea(AreaLetra, AreaVal, CantAreas, AreaMax, compCentral, compCompart, compTecPorc, compCompPorc)
Call DatosResuObj(objsVal, cantObjs, totalObjs)
Call DatosObjPorcentaje(totalObjs, totalObjGral, ObjsPorc, ObjGralPorc)

     ' los porcentajes compTecPorc, compCompPorc  estan en la escala 1 a 100, se pueden usar como resultados!
     totalComp = (compTecPorc * porcCTec / 100 + compCompPorc * porcCComp / 100) ' compCentral + compCompart
     
     
     ' los porcentajes ObjsPorc, ObjGralPorc,CalifGralRDPPorc, estan en la escala 1 a 100, se pueden usar como resultados!
     Select Case tipoEv
     Case 1:
         totalPuntos = (compTecPorc * porcCTec / 100 + compCompPorc * porcCComp / 100 + ObjsPorc * porcGral / 100)     ' totalObjs     ' prom - calif Indiv de Objs
     Case 2:
         totalPuntos = (compTecPorc * porcCTec / 100 + compCompPorc * porcCComp / 100 + ObjGralPorc * porcGral / 100) ' totalObjGral  ' prom - calif gral objs
     Case 3:
         totalPuntos = (compTecPorc * porcCTec / 100 + compCompPorc * porcCComp / 100 + CalifGralRDPPorc * porcGral / 100) 'CalifGralRDP o Ev. Gral  ' prom - calficac gral de la rdp
     Case Else
         totalPuntos = (compTecPorc * porcCTec / 100 + compCompPorc * porcCComp / 100 + CalifGralRDPPorc * porcGral / 100) ' CalifGralRDP - o Ev. Gral
     End Select
               
Call DatosEvaFinal(totalPuntos, evalFinal)
Call DatosCalifPreliminar(totalPuntos, califPrel)
Call DatosTipoEvGral(Ternro, tipoEvGral, resuEvaGral)


' evalfinal, totalcomp

Campos = " (bpronro,Fecha,Hora,repnro ,Evento, tipoev   "    'Fecha - por ahora la del proceso - vER
Valores = "(" & NroProceso & "," & ConvFecha(bpfecha) & ",'" & bphora & "', " & repNro & "," & Evento & "," & tipoEv

'  porcentajes asignados a Competencia Tecnicas, Compartidas y a Ev Gral u  Objetivos
Campos = Campos & " , porcCComp, porcCTec, porcGral "
Valores = Valores & "," & numberForSQL(porcCComp) & ", " & numberForSQL(porcCTec) & ", " & numberForSQL(porcGral)

Campos = Campos & ", empleg, ternro,apeynom,edad, fnac, consejero "
Valores = Valores & "," & empleg & "," & Ternro & ",'" & ApeyNom & "'"
Valores = Valores & "," & Edad & "," & ConvFecha(FecNac)
Valores = Valores & ",'" & ConsejeroDesc & "'"

Campos = Campos & ",grupo, depto,categ,categ_estrnro,categ_fecha, antcat, antnum, convenio,empant "  'EmpEstrnro1 (tenro1, estrnro1)
Valores = Valores & ",'" & Estrnro1Desc & "'"  'Grupo
Valores = Valores & ",'" & Estrnro2Desc & "'"  'Depto
Valores = Valores & ",'" & Estrnro3Desc & "'"  'Categ
Valores = Valores & "," & Categ_Estrnro 'estrnro de la categoria
Valores = Valores & "," & ConvFecha(Categ_fecha) 'antiguedad de la categoria
Valores = Valores & ",'" & AntigCat & "'"
Valores = Valores & "," & AntigNum
Valores = Valores & ",' " & Convenio & "'"
Valores = Valores & ",' " & EmpAnt & "'"

  ' Datos del filtro
Campos = Campos & ",todosgrupo,todosdepto,todoscateg,todosconsej, todosaconsej "
Valores = Valores & "," & TodosGrupo
Valores = Valores & "," & TodosDepto
Valores = Valores & "," & TodosCateg
Valores = Valores & "," & TodosConsej
Valores = Valores & "," & TodosAconsej

' AREAS ------------------------------------------------------
Campos = Campos & " ,area1,area2, area3,area4, area5, arean  "
  ' valores de las areas (fijas)
For i = 1 To 5
     Valores = Valores & ",'" & AreaLetra(i) & "' "
Next
  ' valores de las areas restantes
Valores = Valores & ",'0"
For i = 6 To CantAreas
     Valores = Valores & "@" & Area(i) & "-" & AreaLetra(i) & "-" & AreaVal(i)
Next
Valores = Valores & "'"
   
' RDP o Ev. Gral ------------------
Campos = Campos & ", califrdp, califgralRDP "
Valores = Valores & ", '" & CalifRDP & "' " & ", " & numberForSQL(CalifGralRDPPorc)   ' numberForSQL(CalifGralRDP)

' OBJETIVOS ----------------------------------------------------------------
Campos = Campos & " ,calif1, calif2, calif3,calif4,calif5,calif6,cantobj "
For i = 1 To 6
     Valores = Valores & "," & ObjsCant(i)
Next
Valores = Valores & "," & cantObjs

Campos = Campos & " , cantobjgral, totalobjgral "
Valores = Valores & "," & cantObjGral & ", " & numberForSQL(ObjGralPorc)  ' porc escala 1-100, numberForSQL(totalObjGral)

' Totales de AREAS y OBJS
Campos = Campos & " ,compcentral, compcomp, totalobj,evalfinal, totalcomp  "
Valores = Valores & "," & numberForSQL(compTecPorc) & "," & numberForSQL(compCompPorc)    '  numberForSQL(compCentral) & "," & numberForSQL(compCompart)
Valores = Valores & "," & numberForSQL(ObjsPorc) ' numberForSQL(totalObjs)
Valores = Valores & ",'" & evalFinal & "'," & numberForSQL(totalComp)

' Datos para Pociciones - Curva - Ajuste , etc
Campos = Campos & " ,evalgral ,sueldo "
Valores = Valores & "," & numberForSQL(totalPuntos) & "," & numberForSQL(sueldo)

  ' Calif prelim - desempeño - calific ajustada
Campos = Campos & " ,califprel, califajus "
Valores = Valores & "," & califPrel & "," & califPrel

' tipoEvGral, resuEvaGral
Campos = Campos & " ,tipoevgral, resutevagral "
Valores = Valores & ", '" & tipoEvGral & "' , '" & resuEvaGral & "' "

Campos = Campos & ")"
Valores = Valores & ")"
 

StrSql = " INSERT INTO rep_stulich " & Campos & " VALUES " & Valores
objConn.Execute StrSql, , adExecuteNoRecords


' --------------
'    'Cierro y libero
'    If rsConsult1.State = adStateOpen Then rsConsult1.Close
'    If rsConsult2.State = adStateOpen Then rsConsult2.Close
'    Set rsConsult1 = Nothing
'    Set rsConsult2 = Nothing
Exit Sub

MError:
    Flog.writeline "    Error - en el Tercero: " & Ternro
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    Resume Next
    HuboErrores = True
    EmpErrores = True
    Exit Sub
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



' ------------------------------------------------------------
'
' ------------------------------------------------------------
Sub DatosEstructura(Ternro, ByRef EmpEstrnro1, ByRef EmpEstrnro2, ByRef EmpEstrnro3, ByRef Estrnro1Desc, ByRef Estrnro2Desc, ByRef Estrnro3Desc)

On Error GoTo ME_Estr

'Busco los datos del tipos de estructura 1
If Tenro1 <> 0 Then
    If Estrnro1 <> 0 And Estrnro1 <> -1 Then
        EmpEstrnro1 = Estrnro1
        StrSql = " SELECT estrdabr, estrnro FROM estructura  WHERE estrnro = " & Estrnro1 & " AND tenro = " & Tenro1
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
             Estrnro1Desc = rsConsult!estrdabr
        End If
    Else
        StrSql = " SELECT his_estructura.estrnro, estrdabr FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = " & Tenro1
        StrSql = StrSql & " AND (htethasta is null )"
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro1 = rsConsult!estrnro
           Estrnro1Desc = rsConsult!estrdabr
        End If
    End If
End If

'Busco los datos del tipos de estructura 2
If Tenro2 <> 0 Then
    If Estrnro2 <> 0 And Estrnro2 <> -1 Then
        EmpEstrnro2 = Estrnro2
        StrSql = " SELECT estrdabr, estrnro FROM estructura  WHERE estrnro = " & Estrnro2 & " AND tenro = " & Tenro2
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
             Estrnro2Desc = rsConsult!estrdabr
        End If
    Else
        StrSql = " SELECT his_estructura.estrnro, estrdabr FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = " & Tenro2
        StrSql = StrSql & " AND (htethasta is null )"
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro2 = rsConsult!estrnro
           Estrnro2Desc = rsConsult!estrdabr
        End If
    End If
End If

'Busco los datos del tipos de estructura 3
If Tenro3 <> 0 Then
    If Estrnro3 <> 0 And Estrnro3 <> -1 Then
        EmpEstrnro3 = Estrnro3
        StrSql = " SELECT estrdabr, estrnro FROM estructura  WHERE estrnro = " & Estrnro3 & " AND tenro = " & Tenro3
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
             Estrnro3Desc = rsConsult!estrdabr
             Categ_Estrnro = rsConsult!estrnro
        End If
    Else
        StrSql = " SELECT his_estructura.estrnro, estrdabr FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = " & Tenro3
        StrSql = StrSql & " AND (htethasta is null )"
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro3 = rsConsult!estrnro
           Estrnro3Desc = rsConsult!estrdabr
           Categ_Estrnro = rsConsult!estrnro
        End If
    End If
End If

Exit Sub

ME_Estr:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub

' ------------------------------------------------------------
' Busca los datos del empleado
' ------------------------------------------------------------
Sub DatosEmp(Ternro, ByRef empleg, ByRef ApeyNom, ByRef FecNac, ByRef Edad)

On Error GoTo ME_Emp

StrSql = "SELECT tercero.ternom, tercero.ternom2, tercero.terape,tercero.terape2, tercero.terfecnac, empleado.empleg "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " INNER JOIN tercero ON empleado.ternro = tercero.ternro "
StrSql = StrSql & " WHERE tercero.ternro =" & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    empleg = rsConsult!empleg
    ApeyNom = rsConsult!terape & " " & rsConsult!terape2 & " " & rsConsult!ternom & " " & rsConsult!ternom2
    FecNac = rsConsult!terfecnac
    'empleg ApeyNom FecNac
End If
rsConsult.Close

' busca la edad ....
If (Month(bpfecha) > Month(CDate(FecNac))) Then
     Edad = DateDiff("yyyy", CDate(FecNac), bpfecha)
Else
     If (Month(bpfecha) = Month(CDate(FecNac))) And (Day(bpfecha) > Day(CDate(FecNac))) Then
        Edad = DateDiff("yyyy", CDate(FecNac), bpfecha)
     Else
        Edad = DateDiff("yyyy", CDate(FecNac), bpfecha) - 1
     End If
End If

Exit Sub

ME_Emp:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub

' _______________________________________________________________
'  Busca si tiene convenio o no - y si tenia empleo anterior
' _______________________________________________________________
Sub DatosVarios(Ternro, ByRef Convenio, ByRef EmpAnt)

On Error GoTo ME_varios

StrSql = " SELECT his_estructura.estrnro, estrdabr FROM his_estructura "
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
'StrSql = StrSql & " INNER JOIN convenios ON convenios.estrnro = his_estructura.estrnro "
StrSql = StrSql & " WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = 19 " ' tenro=19 - convenio
StrSql = StrSql & " AND (htethasta is null )"
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    If (UCase(Trim(rsConsult!estrdabr)) = "SIN CONVENIO" Or UCase(Trim(rsConsult!estrdabr)) = "FUERA DE CONVENIO") Then
        Convenio = "NO"
    Else
        Convenio = "SI"
    End If
Else
    Convenio = "NO"
End If


StrSql = " SELECT empleado FROM empant WHERE empleado = " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    EmpAnt = "ex AA"
Else
    EmpAnt = "Deloitte"
End If

rsConsult.Close

Exit Sub

ME_varios:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub

' ------------------------------------------------------------
'  Busca los datos del consejero
' ------------------------------------------------------------
Sub DatosConsejero(Ternro, Consejero, ByRef ConsejeroDesc)
Dim consejAux
On Error GoTo ME_consej

consejAux = Consejero

If consejAux = 0 Then
   StrSql = "SELECT DISTINCT evadetevldor.evaluador,  evaoblieva.evaobliorden "
   StrSql = StrSql & " FROM evacab "
   StrSql = StrSql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
   StrSql = StrSql & " INNER JOIN evaoblieva ON evaoblieva.evatevnro = evadetevldor.evatevnro AND evaoblieva.evaseccnro = evadetevldor.evaseccnro "
   StrSql = StrSql & " WHERE evacab.empleado =" & Ternro & " AND evacab.evaevenro = " & Evento
   StrSql = StrSql & " ORDER BY evaoblieva.evaobliorden DESC "
   ' & " AND evadetevldor.evaluador <> " & Ternro - evadetevldor.evatevnro,
   OpenRecordset StrSql, rsConsult

   If Not rsConsult.EOF Then
       consejAux = IIf(Not EsNulo(rsConsult!evaluador), rsConsult!evaluador, 0)
   End If
End If

StrSql = "SELECT tercero.ternom, tercero.ternom2, tercero.terape,tercero.terape2, tercero.terfecnac, empleado.empleg "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " INNER JOIN tercero ON empleado.ternro = tercero.ternro "
StrSql = StrSql & " WHERE tercero.ternro =" & consejAux
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    ConsejeroDesc = rsConsult!empleg & " - " & rsConsult!terape & " " & rsConsult!terape2 & " " & rsConsult!ternom & " " & rsConsult!ternom2
End If
rsConsult.Close

Exit Sub

ME_consej:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub

' VERRRRRRR el que esta en el modulo va sumando la cant dias en las distintas fases --> esto si es antig en toda la empresa!
Public Sub bus_AntEstructura(Ternro, EmpEstrnro3, Categ_fecha, ByRef AntigCat, ByRef AntigNum As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Antiguedad en la Estructura a una Fecha
' Autor      : FGZ
' Fecha      : 25/11/2003
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim TipoEstr As Long      ' Tipo de Estructura
Dim TipoFecha As Integer    ' 1 - Primer dia del año
                            ' 2 - Ultimo dia del año
                            ' 3 - Inicio del proceso
                            ' 4 - Fin del proceso
                            ' 5 - Inicio del Periodo
                            ' 6 - Fin del periodo
                            ' 7 - Today
Dim Resultado As Integer    ' Tipo de resultado devuelto
                            ' 1 - En dias
                            ' 2 - En Meses
                            ' 3 - En Años
                            
'Dim Param_cur As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset

Dim Aux_Fecha As Date

Dim FechaDesde As Date
Dim FechaHasta As Date

Dim Dia As Integer
Dim mes As Integer
Dim anio As Integer
   
Dim aux1 As Integer
Dim aux2 As Integer
Dim aux3 As Integer

Dim Aux_Meses As Integer

On Error GoTo ME_antEstr

    'If Arr_Programa(NroProg).Prognro <> 0 Then
        'TipoEstr = Arr_Programa(NroProg).Auxint1
        'TipoFecha = Arr_Programa(NroProg).Auxint2
        'Resultado = Arr_Programa(NroProg).Auxint3
    'Else
        'Exit Sub
    'End If
    
    TipoEstr = Tenro3  ' Categoria
    TipoFecha = 8 ' la Antiguedad en la Categoria
Dia = 0
mes = 0
anio = 0

Select Case TipoFecha
Case 1: 'Primer dia del año
    Aux_Fecha = CDate("01/01/" & Year(Date))
Case 2: 'Ultimo dia del año
    Aux_Fecha = CDate("31/12/" & Year(Date))
Case 3: 'Inicio del proceso
    'Aux_Fecha = buliq_proceso!profecini
Case 4: 'Fin del proceso
    'Aux_Fecha = buliq_proceso!profecfin
Case 5: 'Inicio del periodo
    'Aux_Fecha = buliq_periodo!pliqdesde
Case 6: 'Fin del periodo
    'Aux_Fecha = buliq_periodo!pliqhasta
Case 7: 'Today
    Aux_Fecha = bpfecha ' Date
Case 8:
    Aux_Fecha = Categ_fecha
Case Else
    'tipo de fecha no valido
End Select


'FGZ - 19/03/2004
' XXXXXXX antes de usarlo tengo que obtener empest
'If Not CBool(buliq_empleado!empest) Then
    'If Aux_Fecha > Empleado_Fecha_Fin Then
        'Aux_Fecha = Empleado_Fecha_Fin
    'End If
'End If

If Not EsNulo(Aux_Fecha) Then
    ' Busco de estructura '
        StrSql = " SELECT htetdesde,htethasta FROM his_estructura " & _
                 " WHERE ternro = " & Ternro & _
                 "   AND tenro =" & TipoEstr & " AND estrnro=" & EmpEstrnro3 & _
                 "   AND (htetdesde <= " & ConvFecha(Aux_Fecha) & ") AND " & _
                 "   ((" & ConvFecha(Aux_Fecha) & " <= htethasta) or (htethasta is null))"
        OpenRecordset StrSql, rs_Estructura
    
        If Not rs_Estructura.EOF Then
            FechaDesde = rs_Estructura!htetdesde
            FechaHasta = IIf(EsNulo(rs_Estructura!htethasta), Date, rs_Estructura!htethasta)
        End If
        
        Call Dif_Fechas(FechaDesde, Aux_Fecha, aux1, aux2, aux3)
        Dia = Dia + aux1
        'mes = mes + aux2 + Int(dia / 30)
        'anio = anio + aux3 + Int(mes / 12)
        'dia = dia Mod 30
        'mes = mes Mod 12
        
    ' -- codigo en base a antigfec() - en .asp
    If Dia > 0 Then
        ' calcular año
         If Dia > 365 Then
            anio = Int(Dia / 365)
            Dia = Dia - 365 * anio + 1
         End If
         If Dia > 30 Then
            mes = Int(Dia / 30.5)
            If Month(Aux_Fecha) = 4 Or Month(Aux_Fecha) = 6 Or Month(Aux_Fecha) = 9 Or Month(Aux_Fecha) = 11 Then
                Dia = Dia - Int(mes * 30.5) + 1
            Else
                If Month(Aux_Fecha) = 2 Then
                    Dia = Dia - Int(mes * 28) + 1
                Else
                    Dia = Dia - Int(mes * 30.5) + 1
                End If
            End If
         Else
            Dia = Dia + 1
         End If
    End If

    'Select Case Resultado
    'Case 1: ' En dias       'Valor = dia
    'Case 2: ' En meses      'Valor = mes
    'Case 3: ' en años       'Valor = anio
    'End Select
End If
    
AntigCat = ""
If anio <> 0 Then
    If anio = 1 Then
        AntigCat = anio & " año "
    Else
        AntigCat = anio & " años "
    End If
End If

If mes = 0 Then
    'AntigCat = Dia & " día/s."
Else
    If mes = 1 Then
        AntigCat = AntigCat & mes & " mes "
    Else
        AntigCat = AntigCat & mes & " meses "
    End If
End If


Aux_Meses = anio * 12 + mes
Select Case Aux_Meses
Case Is <= 7:
    AntigNum = 1
Case Is <= 19:
    AntigNum = 2
Case Is <= 31:
    AntigNum = 3
Case Is <= 43:
    AntigNum = 4
Case Is <= 55:
    AntigNum = 5
Case Else
    AntigNum = 6
End Select

If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
Set rs_Estructura = Nothing

Exit Sub

ME_antEstr:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub
' _______________________________________________________________
' busca el sueldo del empleado - Usa periodod del evento (fecha hasta)
' por ahora se fija solo en ACUmulador !!!!!!!!11
' _______________________________________________________________
Sub DatosSueldo(Ternro, Acu, CO, ByRef sueldo)
Dim FHasta As Date
Dim mes
Dim anio
On Error GoTo ME_sueldo

FHasta = Date
mes = 0
anio = 0
sueldo = 0

'StrSql = " SELECT evaevefdesde, evaevefhasta FROM evaevento WHERE evaevenro =" & Evento
'OpenRecordset StrSql, rsConsult
'If Not rsConsult.EOF Then
    'FHasta = rsConsult!evaevefhasta
'End If

mes = Month(FHasta)
anio = Year(FHasta)

If mes = 1 Then
    mes = 12
    anio = anio - 1
Else
    mes = mes - 1
End If

StrSql = " SELECT ammonto FROM acu_mes "
StrSql = StrSql & " WHERE amanio =" & anio & " AND ammes=" & mes
StrSql = StrSql & "   AND ternro=" & Ternro & " AND acunro=" & Acu

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    sueldo = rsConsult!ammonto
End If

rsConsult.Close

Exit Sub

ME_sueldo:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub

' ______________________________________________________________________________
' busco un solo resultado MAximo para las Competencias u Areas (se supone que todas las areas y competencias tienen el mismo conjunto de resultados)
' ______________________________________________________________________________
Sub buscarMaxValorArea(ByRef maxValorArea)
Dim rsConsult1  As New ADODB.Recordset
Dim StrSql As String

    maxValorArea = 0

    StrSql = "SELECT  evaresu.evatrnro, evatipresu.evatrvalor "
    StrSql = StrSql & " FROM evaresu "
    StrSql = StrSql & " INNER JOIN evatipresu ON evatipresu.evatrnro = evaresu.evatrnro "
    StrSql = StrSql & " INNER JOIN evasecc ON evasecc.evaseccnro= evaresu.evaseccnro "
    StrSql = StrSql & " INNER JOIN evatipoeva ON evatipoeva.evatipnro = evasecc.evatipnro "
    StrSql = StrSql & " INNER JOIN evaevento ON evaevento.evatipnro= evatipoeva.evatipnro "
    StrSql = StrSql & " WHERE evaevento.evaevenro =" & Evento
    StrSql = StrSql & " ORDER BY evatipresu.evatrvalor DESC "
    OpenRecordset StrSql, rsConsult

    If Not rsConsult.EOF Then
        maxValorArea = rsConsult!evatrvalor
    End If
    
    rsConsult.Close
    
    ' VERRRRRRRR
    ' StrSql = StrSql & " INNER JOIN evaseccgral ON evaseccgral.evatgralnro = evagralposresu.evatgralnro  AND evaseccgral.evaseccnro = evagralposresu.evaseccnro "

End Sub


' ----------------------------------------------------------------
'  resultado de la Evaluacion sobre las Areas
' ----------------------------------------------------------------
Sub DatosArea(Ternro, ByRef Area, ByRef AreaLetra, ByRef AreaVal, ByRef CantAreas, ByRef AreaMax)
Dim i
Dim j
Dim rsConsult1  As New ADODB.Recordset
Dim StrSql1 As String

Dim maxValorArea As Integer


On Error GoTo ME_Areas

For i = 1 To 90
    Area(i) = ""
    AreaLetra(i) = ""
    AreaVal(i) = 0
    AreaMax(i) = 0
Next

buscarMaxValorArea maxValorArea

StrSql = " SELECT DISTINCT evatitulo.evatitnro, evatitulo.evatitdesabr " & _
         " FROM evaevento " & _
         " INNER JOIN evatipoeva ON evatipoeva.evatipnro = evaevento.evatipnro " & _
         " INNER JOIN evasecc ON evasecc.evatipnro = evatipoeva.evatipnro " & _
         " INNER JOIN evaseccfactor ON evaseccfactor.evaseccnro= evasecc.evaseccnro " & _
         " INNER JOIN evafactor ON evafactor.evafacnro = evaseccfactor.evafacnro " & _
         " INNER JOIN evatitulo ON evatitulo.evatitnro = evafactor.evatitnro " & _
         " WHERE evaevento.evaevenro =" & Evento & _
         " ORDER BY evatitulo.evatitnro "
        ' l_sql = " INNER JOIN evacab ON evacab.evaevento = "
OpenRecordset StrSql, rsConsult


' si da vacio, signific que la seccion es solamente de Areas (y no Areas y Competencias), busco en config de seccareas
If rsConsult.EOF Then
    rsConsult.Close
    
    StrSql = " SELECT DISTINCT evatitulo.evatitnro, evatitulo.evatitdesabr " & _
             " FROM evaevento " & _
             " INNER JOIN evatipoeva ON evatipoeva.evatipnro = evaevento.evatipnro " & _
             " INNER JOIN evasecc ON evasecc.evatipnro = evatipoeva.evatipnro " & _
             " INNER JOIN evaseccarea ON evaseccarea.evaseccnro= evasecc.evaseccnro " & _
             " INNER JOIN evatitulo ON evatitulo.evatitnro = evaseccarea.evatitnro " & _
             " WHERE evaevento.evaevenro =" & Evento & _
             " ORDER BY evatitulo.evatitnro "
            ' l_sql = " INNER JOIN evacab ON evacab.evaevento = "
    OpenRecordset StrSql, rsConsult
End If


i = 6
j = 6
CantAreas = 5   ' Comp Compartidas y Tecnicas Centrales

Do While Not rsConsult.EOF  ' en BBca 10 - 11 - 1 - 2 - 12

    Select Case rsConsult!evatitnro
        Case 23:
                Area(1) = rsConsult!evatitnro
                j = 1
        Case 24:
                Area(2) = rsConsult!evatitnro
                j = 2
        Case 25:
                Area(3) = rsConsult!evatitnro
                j = 3
        Case 26:
                Area(4) = rsConsult!evatitnro
                j = 4
        Case 27:
                Area(5) = rsConsult!evatitnro
                j = 5
        Case Else:
                  Area(i) = rsConsult!evatitnro
                  j = i
    End Select
    
    
    StrSql1 = " SELECT evacab.evacabnro, evadetevldor.evldrnro, evaarea.evatitnro, evaarea.evatrnro, evatrletra, evatrvalor "
    StrSql1 = StrSql1 & " FROM evacab "
    StrSql1 = StrSql1 & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
    StrSql1 = StrSql1 & " INNER JOIN evaarea ON evaarea.evldrnro =evadetevldor.evldrnro "
    StrSql1 = StrSql1 & " LEFT  JOIN evatipresu ON evatipresu.evatrnro = evaarea.evatrnro "
    StrSql1 = StrSql1 & " WHERE evacab.evaevenro=" & Evento
    StrSql1 = StrSql1 & "   AND evacab.empleado=" & Ternro
    StrSql1 = StrSql1 & "   AND evaarea.evatitnro=" & rsConsult!evatitnro
    StrSql1 = StrSql1 & "   AND evadetevldor.evatevnro=" & cconsejero
    OpenRecordset StrSql1, rsConsult1
    
    If Not rsConsult1.EOF Then
        If Not EsNulo(rsConsult1!evatrnro) Then
            AreaLetra(j) = rsConsult1!evatrletra
            AreaVal(j) = rsConsult1!evatrvalor
        End If
    End If
    rsConsult1.Close
    
    ' ___________________________
    AreaMax(j) = maxValorArea
    
    
       '  B. Bca. ' en BBca 10 - 11 - 1 - 2 - 12
       '  If rsConsult!evatitnro <> 10 And rsConsult!evatitnro <> 11 And rsConsult!evatitnro <> 1 And rsConsult!evatitnro <> 2 And rsConsult!evatitnro <> 12 Then
    If rsConsult!evatitnro <> 23 And rsConsult!evatitnro <> 24 And rsConsult!evatitnro <> 25 And rsConsult!evatitnro <> 26 And rsConsult!evatitnro <> 27 Then
        i = i + 1
        CantAreas = CantAreas + 1
    End If
    
rsConsult.MoveNext
Loop

rsConsult.Close



Exit Sub

ME_Areas:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub

' _______________________________________________________________
'  Calificacion Gral de la RDP.
' 16-05-2008 - buscar Calificacion Final de RDP o Evaluacion Final segun Formulario
' _______________________________________________________________
Sub DatosRDP(Ternro, ByRef CalifRDP, ByRef CalifGralRDP, ByRef CalifGralRDPPorc)
Dim maxValorGralRDP
Dim hayCalifGralRDP

On Error GoTo ME_RDP

CalifRDP = ""
CalifGralRDP = 0
maxValorGralRDP = 0
CalifGralRDPPorc = 0
hayCalifGralRDP = 1

' me fijo si existe Ev General con tipo ev gral con orden -1
StrSql = "SELECT  evagralposresu.evatrnro, evatipresu.evatrvalor "
StrSql = StrSql & " FROM evagralposresu "
StrSql = StrSql & " INNER JOIN evatipresu ON evatipresu.evatrnro = evagralposresu.evatrnro "
StrSql = StrSql & " INNER JOIN evaseccgral ON evaseccgral.evatgralnro = evagralposresu.evatgralnro  AND evaseccgral.evaseccnro = evagralposresu.evaseccnro "
StrSql = StrSql & " INNER JOIN evasecc ON evasecc.evaseccnro = evaseccgral.evaseccnro "
StrSql = StrSql & " INNER JOIN evatipoeva ON evatipoeva.evatipnro = evasecc.evatipnro "
StrSql = StrSql & " INNER JOIN evaevento ON evaevento.evatipnro= evatipoeva.evatipnro "
StrSql = StrSql & " WHERE evaevento.evaevenro =" & Evento
StrSql = StrSql & "  AND  evaseccgral.orden = -1 "
StrSql = StrSql & " ORDER BY evatipresu.evatrvalor DESC "
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    maxValorGralRDP = rsConsult!evatrvalor
    hayCalifGralRDP = 0
End If
rsConsult.Close


If hayCalifGralRDP = 0 Then
    StrSql = " SELECT evacab.evacabnro, evadetevldor.evldrnro, evagralresu.evatrnro, evatrletra, evatrvalor "
    StrSql = StrSql & " FROM evacab "
    StrSql = StrSql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
    StrSql = StrSql & " INNER JOIN evagralresu ON evagralresu.evldrnro =evadetevldor.evldrnro "
    StrSql = StrSql & " INNER JOIN evaseccgral ON evaseccgral.evatgralnro = evagralresu.evatgralnro  AND evaseccgral.evaseccnro = evadetevldor.evaseccnro "
    StrSql = StrSql & " LEFT  JOIN evatipresu ON evatipresu.evatrnro = evagralresu.evatrnro "
    StrSql = StrSql & " WHERE evacab.evaevenro=" & Evento
    StrSql = StrSql & "     AND evacab.empleado=" & Ternro
    StrSql = StrSql & "     AND  evaseccgral.orden = -1 "
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        CalifRDP = rsConsult!evatrletra
        If Not EsNulo(rsConsult!evatrnro) Then
            If rsConsult!evatrletra <> "" And (UCase(rsConsult!evatrletra) <> "NA" And UCase(rsConsult!evatrletra) <> "NA.") Then
            CalifGralRDP = rsConsult!evatrvalor
            End If
        End If
    End If
    rsConsult.Close
    
    If maxValorGralRDP <> 0 Then
        CalifGralRDPPorc = (CalifGralRDP * 100) / maxValorGralRDP
    End If
    
Else

    StrSql = " SELECT evacab.evacabnro, evadetevldor.evldrnro, evagralrdp.evatrnro, evatrletra, evatrvalor "
    StrSql = StrSql & " FROM evacab "
    StrSql = StrSql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
    StrSql = StrSql & " INNER JOIN evagralrdp ON evagralrdp.evldrnro = evadetevldor.evldrnro "
    StrSql = StrSql & " LEFT  JOIN evatipresu ON evatipresu.evatrnro = evagralrdp.evatrnro "
    StrSql = StrSql & " WHERE evacab.evaevenro=" & Evento & " AND evacab.empleado=" & Ternro
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        CalifRDP = rsConsult!evatrletra
        If Not EsNulo(rsConsult!evatrnro) Then
            If rsConsult!evatrletra <> "" And (UCase(rsConsult!evatrletra) <> "NA" And UCase(rsConsult!evatrletra) <> "NA.") Then
            CalifGralRDP = rsConsult!evatrvalor  ' valorRDP(rsConsult!evatrletra)
            End If
        End If
    End If
    rsConsult.Close


    maxValorGralRDP = 0
    
    StrSql = "SELECT  evatipresu.evatrnro, evatipresu.evatrvalor "
    StrSql = StrSql & " FROM evatipresu  "
    StrSql = StrSql & " WHERE evatrtipo=1 "
    StrSql = StrSql & " ORDER BY evatipresu.evatrvalor DESC "
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
         maxValorGralRDP = rsConsult!evatrvalor
    End If
    rsConsult.Close
    
    If maxValorGralRDP <> 0 Then
        CalifGralRDPPorc = (CalifGralRDP * 100) / maxValorGralRDP
    End If
    
End If



Exit Sub

ME_RDP:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub


' -----------------------------------------------------------
'  Busca la cantidad de Objetivos que tiene cada Calificac
' -----------------------------------------------------------
Sub DatosObj(Ternro, ByRef ObjsCant, ByRef objsVal, ByRef cantObjs)
Dim i
Dim StrSql1 As String
Dim rsConsult1 As New ADODB.Recordset
On Error GoTo ME_Obj

StrSql = " SELECT evatrnro, evatrdesabr, evatrvalor, evatrletra  FROM evatipresu WHERE evatrtipo=2 ORDER BY evatrnro"
OpenRecordset StrSql, rsConsult

i = 1
cantObjs = 0

Do While Not rsConsult.EOF
    
    StrSql1 = "SELECT DISTINCT evaobjetivo.evaobjnro "
    StrSql1 = StrSql1 & " FROM evaevento "
    StrSql1 = StrSql1 & " INNER JOIN evaproyecto ON evaproyecto.evapernro = evaevento.evaperact "
    StrSql1 = StrSql1 & " INNER JOIN evatipoeva  ON evatipoeva.evatipnro = evaevento.evatipnro "
    StrSql1 = StrSql1 & " INNER JOIN evatip_estr ON evatip_estr.evatipnro = evatipoeva.evatipnro AND evatip_estr.estrnro=evaproyecto.estrnro "
    StrSql1 = StrSql1 & " INNER JOIN evacab ON evacab.evaproynro = evaproyecto.evaproynro "
    StrSql1 = StrSql1 & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
    StrSql1 = StrSql1 & " INNER JOIN evaluaobj ON evaluaobj.evldrnro = evadetevldor.evldrnro "
    StrSql1 = StrSql1 & " INNER JOIN evaobjetivo ON evaobjetivo.evaobjnro = evaluaobj.evaobjnro "
        ' los unicos que van a tener valor, evatrnro,  son lo de la seccion Calific Objs ( Def objs y Notas Objs --> NO)
    StrSql1 = StrSql1 & " INNER JOIN evatipresu ON evatipresu.evatrnro = evaluaobj.evatrnro "
    StrSql1 = StrSql1 & " WHERE evaevento.evaevenro =" & Evento & " AND evacab.empleado =" & Ternro & " AND evadetevldor.evatevnro =" & cevaluador & " AND evaluaobj.evatrnro=" & rsConsult!evatrnro
    OpenRecordset StrSql1, rsConsult1
    
    ObjsCant(i) = 0
    objsVal(i) = 0
    'If Not rsConsult1.EOF Then
        Do While Not rsConsult1.EOF
            ObjsCant(i) = ObjsCant(i) + 1
               ' Cant - --
            If rsConsult!evatrletra <> "" And (UCase(rsConsult!evatrletra) <> "NA" And UCase(rsConsult!evatrletra) <> "NA.") Then
                objsVal(i) = objsVal(i) + CDbl(rsConsult!evatrvalor)
                cantObjs = cantObjs + 1
            End If
        rsConsult1.MoveNext
        Loop
    'Else
        'ObjsCant(i) = 0
        'ObjsVal(i) = 0
    'End If
    rsConsult1.Close
    
    i = i + 1
rsConsult.MoveNext
Loop

rsConsult.Close

Exit Sub

ME_Obj:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub


' _______________________________________________________________
' Saca un promedio de La calificacion de Obj Gral de las RDEs
' _______________________________________________________________

Sub DatosObjGral(Ternro, ByRef promObjGral, ByRef cantObjGral)
Dim suma
On Error GoTo ME_ObjGral

suma = 0
cantObjGral = 0
promObjGral = 0

StrSql = "SELECT DISTINCT evagralobj.evldrnro, evaproyecto.evaproynro, evatipresu.evatrvalor, evatipresu.evatrletra  "
StrSql = StrSql & " FROM evaevento "
StrSql = StrSql & " INNER JOIN evaproyecto ON evaproyecto.evapernro = evaevento.evaperact "
StrSql = StrSql & " INNER JOIN evatipoeva  ON evatipoeva.evatipnro = evaevento.evatipnro "
StrSql = StrSql & " INNER JOIN evatip_estr ON evatip_estr.evatipnro = evatipoeva.evatipnro AND evatip_estr.estrnro=evaproyecto.estrnro "
StrSql = StrSql & " INNER JOIN evacab ON evacab.evaproynro = evaproyecto.evaproynro "
StrSql = StrSql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
StrSql = StrSql & " INNER JOIN evagralobj ON evagralobj.evldrnro = evadetevldor.evldrnro "
StrSql = StrSql & " INNER JOIN evatipresu ON evatipresu.evatrnro = evagralobj.evatrnro "
StrSql = StrSql & " WHERE evaevento.evaevenro =" & Evento
StrSql = StrSql & " AND evacab.empleado =" & Ternro
StrSql = StrSql & " AND evadetevldor.evatevnro =" & cevaluador
OpenRecordset StrSql, rsConsult

Do While Not rsConsult.EOF
   If rsConsult!evatrletra <> "" And (UCase(rsConsult!evatrletra) <> "NA" And UCase(rsConsult!evatrletra) <> "NA.") Then
        suma = suma + CDbl(rsConsult!evatrvalor)
        cantObjGral = cantObjGral + 1
   End If
rsConsult.MoveNext
Loop
rsConsult.Close
    
If cantObjGral <> 0 Then
    promObjGral = suma / cantObjGral
End If

Exit Sub

ME_ObjGral:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub


' _____________________________________________________________________
' ... saca resultado de la evaluacion de la areas - tec centrales y compartidas
' _____________________________________________________________________
Sub DatosResuArea(AreaLetra, AreaVal, CantAreas, AreaMax, ByRef compCentral, ByRef compCompart, ByRef compTecPorc, ByRef compCompPorc)
Dim i
Dim cantProm
Dim totalAreas
Dim AreaMaxValor

On Error GoTo ME_resuArea

compCentral = 0
compCompart = 0
compTecPorc = 0
compCompPorc = 0
AreaMaxValor = 0


For i = 1 To CantAreas
    If AreaMax(i) > AreaMaxValor Then
        AreaMaxValor = AreaMax(i)
    End If
Next


 ' asignar valor para las competencias tec. centrales
    cantProm = 0
    totalAreas = 0
    For i = 1 To CantAreas
        If (i <> 2) And (i <> 3) And (i <> 4) And (i <> 5) Then
     
            If AreaLetra(i) <> "" And (UCase(AreaLetra(i)) <> "NA" And UCase(AreaLetra(i)) <> "NA.") Then
                'If AreaVal(i) <> -1 Then
                    cantProm = cantProm + 1
                    totalAreas = totalAreas + AreaVal(i)
                'End If
            End If
            
        End If
    Next

    If cantProm <> 0 Then
        compCentral = totalAreas / cantProm
        If AreaMaxValor <> 0 Then
            compTecPorc = (totalAreas * 100) / (cantProm * AreaMaxValor)
        End If
    End If
    

 ' asignar valor para las competencias compartidas
    cantProm = 0
    totalAreas = 0
    For i = 2 To 5
         If AreaLetra(i) <> "" And (UCase(AreaLetra(i)) <> "NA" And UCase(AreaLetra(i)) <> "NA.") Then
            'If AreaVal(i) <> -1 Then
               cantProm = cantProm + 1
               totalAreas = totalAreas + AreaVal(i)
             'End If
         End If
    Next

    If cantProm <> 0 Then
        compCompart = totalAreas / cantProm
        If AreaMaxValor <> 0 Then
            compCompPorc = (totalAreas * 100) / (cantProm * AreaMaxValor)
        End If
    End If

Exit Sub

ME_resuArea:
    Flog.writeline "    Error: " & Err.Description
    'Flog.Writeline "    SQL Ejecutado: " & StrSql

End Sub


' ______________________________________________________
'  asignar el valor para los Objetivos
' ______________________________________________________
Sub DatosResuObj(objsVal, cantObjs, totalObjs)
Dim i
Dim total

totalObjs = 0
total = 0

For i = 1 To 6
    total = total + objsVal(i)
Next

If cantObjs <> 0 Then
    totalObjs = total / cantObjs
End If

End Sub

' ________________________________________________
' calificac final segun el puntaje de la evalac.
' ________________________________________________

Sub DatosEvaFinal(totalPuntos, ByRef evalFinal)

If totalPuntos > 79 Then
    evalFinal = "ES"
Else
    If totalPuntos > 59 Then
        evalFinal = "E"
    Else
        If totalPuntos > 39 Then
            evalFinal = "A"
        Else
            If totalPuntos > 19 Then
                evalFinal = "AA"
            Else
                evalFinal = "NA"
            End If
        End If
    End If
End If

End Sub

' ______________________________________________________________________
' devuelve el porcentaje  correspondiente a la evaluacion de objetivos.
' ______________________________________________________________________
Sub DatosObjPorcentaje(totalObjs, totalObjGral, ByRef ObjsPorc, ByRef ObjGralPorc)
Dim maxValorObj

On Error GoTo ME_objporc

    maxValorObj = 0
    ObjsPorc = 0
    ObjGralPorc = 0
    
    StrSql = "SELECT  evatipresu.evatrnro, evatipresu.evatrvalor "
    StrSql = StrSql & " FROM evatipresu  "
    StrSql = StrSql & " WHERE evatrtipo=2 "
    StrSql = StrSql & " ORDER BY evatipresu.evatrvalor DESC "
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
         maxValorObj = rsConsult!evatrvalor
    End If
    rsConsult.Close

    If maxValorObj <> 0 Then
        ObjsPorc = (totalObjs * 100) / maxValorObj
        ObjGralPorc = (totalObjGral * 100) / maxValorObj
    End If
    

Exit Sub

ME_objporc:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
End Sub



' _________________________________________________________
' Busca Resultados de las evaluaciones Generales - son 4 por ahora
' _________________________________________________________
Sub DatosTipoEvGral(Ternro, ByRef tipoEvGral, ByRef resuEvaGral)
Dim i
Dim evTipoGral(4)
Dim evTipoGralval(4)



On Error GoTo ME_EvGral

tipoEvGral = ""
resuEvaGral = ""

For i = 1 To 4
    evTipoGral(i) = 0
    evTipoGralval(i) = 0
Next


i = 1

StrSql = " SELECT DISTINCT evaseccgral.evatgralnro, evatgraldabr, evaseccgral.orden,  evasecc.orden "
StrSql = StrSql & " FROM evaseccgral "
StrSql = StrSql & " INNER JOIN evatipogral ON evatipogral.evatgralnro = evaseccgral.evatgralnro "
StrSql = StrSql & " INNER JOIN evasecc ON evasecc.evaseccnro = evaseccgral.evaseccnro "
StrSql = StrSql & " INNER JOIN evaevento ON evaevento.evatipnro = evasecc.evatipnro "
StrSql = StrSql & " WHERE  evaevento.evaevenro=" & Evento
StrSql = StrSql & " ORDER BY evasecc.orden, evaseccgral.orden "
OpenRecordset StrSql, rsConsult

Do While Not rsConsult.EOF
    If i <= 4 Then
    evTipoGral(i) = rsConsult!evatgralnro
    End If
    i = i + 1
rsConsult.MoveNext
Loop
rsConsult.Close


For i = 1 To 4 ' UBound(l_evtipogral)
    
    StrSql = " SELECT DISTINCT evacab.evacabnro, evadetevldor.evldrnro, evagralresu.evatrnro, evatrdesabr, evatrletra, evatrvalor, evagralresu.evatgralnro "
    StrSql = StrSql & " FROM evacab "
    StrSql = StrSql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
    StrSql = StrSql & " INNER JOIN evagralresu ON evagralresu.evldrnro = evadetevldor.evldrnro "
    StrSql = StrSql & " LEFT  JOIN evatipresu ON evatipresu.evatrnro = evagralresu.evatrnro "
    StrSql = StrSql & " WHERE evacab.evaevenro=" & Evento
    StrSql = StrSql & "     AND evacab.empleado=" & Ternro
    StrSql = StrSql & "   AND  evagralresu.evatgralnro=" & evTipoGral(i)
    OpenRecordset StrSql, rsConsult
  
    If Not rsConsult.EOF Then
        If Not IsNull(rsConsult!evatrnro) Then
            evTipoGralval(i) = rsConsult!evatrnro 'rsConsult!evatrvalor
        Else
            evTipoGralval(i) = 0
        End If
    End If
    rsConsult.Close
Next


tipoEvGral = 0
resuEvaGral = 0

For i = 1 To 4
    tipoEvGral = tipoEvGral & "@" & evTipoGral(i)
    resuEvaGral = resuEvaGral & "@" & evTipoGralval(i)
Next

Exit Sub

ME_EvGral:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql


End Sub




' _________________________________________________________
' datos de la calificac preliminar de la evaluac.
' _________________________________________________________
Sub DatosCalifPreliminar(totalPuntos, ByRef califPrel)

califPrel = 1

If totalPuntos < 22.5 Then
    califPrel = 1
Else
    If totalPuntos < 45 Then
        califPrel = 2
    Else
        If totalPuntos < 67.5 Then
            califPrel = 3
        Else
            If totalPuntos < 90 Then
                califPrel = 4
            Else
                califPrel = 5
            End If
        End If
    End If
End If

End Sub


' ________________________________________________________
'  Calificacion Ajustada..  ----
'
' ________________________________________________________
Sub DatosCalifAjustada(bpronro)
Dim Categ
Dim cantEmpls(5)
Dim totalEmpls
Dim cantxporc  'porc,
Dim i

On Error GoTo ME_califAj

Categ = ""

Flog.writeline " "
Flog.writeline " "
Flog.writeline "Se genera la Calificación Ajustada para todos los empleados. "
    
StrSql = " SELECT DISTINCT categ "
StrSql = StrSql & " FROM  rep_stulich "
StrSql = StrSql & " WHERE bpronro =" & bpronro
OpenRecordset StrSql, rsConsult


Do While Not rsConsult.EOF
        
    Categ = rsConsult!Categ
    For i = 1 To 5
        cantEmpls(i) = 0
    Next
    totalEmpls = 0
    
    Call emplsxCalif(bpronro, Categ, cantEmpls, totalEmpls)
    
        ' bajar de puntuacion a empls ...  de menor puntaje...?'
    Call emplsxPorcentaje(totalEmpls, 5, cantxporc)
    If cantEmpls(5) > cantxporc Then
        Call ajustarPuntajeEmpls(bpronro, Categ, 5, cantxporc, cantEmpls)
    End If
            
    Call emplsxPorcentaje(totalEmpls, 15, cantxporc)
    If cantEmpls(4) > cantxporc Then '
        Call ajustarPuntajeEmpls(bpronro, Categ, 4, cantxporc, cantEmpls)
    End If
            
    Call emplsxPorcentaje(totalEmpls, 70, cantxporc)
    If cantEmpls(3) > cantxporc Then
        Call ajustarPuntajeEmpls(bpronro, Categ, 3, cantxporc, cantEmpls)
    End If
    
    rsConsult.MoveNext
Loop

rsConsult.Close

Exit Sub

ME_califAj:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub



'   Cuenta la cant de empleados
    ' _______________________________________________
    Sub emplsxCalif(bpronro, Categ, ByRef cantEmpls, ByRef totalEmpls)
    
    
    Dim rsConsultEmpls As New ADODB.Recordset
        
        StrSql = " SELECT  COUNT(ternro) AS empls, califprel "
        StrSql = StrSql & " FROM  rep_stulich "
        StrSql = StrSql & " WHERE bpronro =" & bpronro & " AND categ= '" & Categ & "'"  ' XXXXXXXXXXXXXX
        StrSql = StrSql & " GROUP BY califprel "
        OpenRecordset StrSql, rsConsultEmpls     'rsOpen rsConsultEmpls, cn, StrSql, 0
        
        Do While Not rsConsultEmpls.EOF
            Select Case rsConsultEmpls!califPrel
                Case 1: cantEmpls(1) = rsConsultEmpls!empls
                Case 2: cantEmpls(2) = rsConsultEmpls!empls
                Case 3: cantEmpls(3) = rsConsultEmpls!empls
                Case 4: cantEmpls(4) = rsConsultEmpls!empls
                Case 5: cantEmpls(5) = rsConsultEmpls!empls
            End Select
            
            totalEmpls = totalEmpls + rsConsultEmpls!empls
        rsConsultEmpls.MoveNext
        Loop
        rsConsultEmpls.Close
    End Sub
    
    ' _______________________________________________
    ' _______________________________________________
    Sub emplsxPorcentaje(totalEmpls, porc, cant)
        cant = 0
        cant = (totalEmpls * porc) / 100
        cant = Round(cant)
    End Sub
     
    ' _____________________________________________________
    ' _____________________________________________________
    Sub ajustarPuntajeEmpls(bpronro, Categ, calif, cantxporc, ByRef cantEmpls) 'xxx
        Dim cant2
        Dim rsConsultEmpls As New ADODB.Recordset
        
        StrSql = " SELECT ternro, evalgral FROM  rep_stulich "
        StrSql = StrSql & " WHERE bpronro =" & bpronro & " AND califprel=" & calif & " AND categ='" & Categ & "' " ' l_rs("categ")
        StrSql = StrSql & " ORDER BY evalgral ASC "
        OpenRecordset StrSql, rsConsultEmpls
         
        cant2 = cantEmpls(calif) - cantxporc
         
        cantEmpls(calif) = cantEmpls(calif) - cant2
        cantEmpls(calif - 1) = cantEmpls(calif - 1) + cant2
         
        Do While Not rsConsultEmpls.EOF And cant2 > 0
            StrSql = " UPDATE rep_stulich SET califajus =" & calif - 1
            StrSql = StrSql & " WHERE bpronro =" & bpronro & "  AND categ='" & Categ & "' AND ternro= " & rsConsultEmpls!Ternro
            objConn.Execute StrSql, , adExecuteNoRecords
             ' l_cm.CommandText = StrSql      cmExecute l_cm, StrSql, 0
            cant2 = cant2 - 1
        rsConsultEmpls.MoveNext
        Loop
        
        rsConsultEmpls.Close
    End Sub
        


Sub AjustarPosicion_Y_Remuneracion()
' ---------------------------------------------------------------------------------------------
' Descripcion: procedimiento que calcula la max remuneracion y la posicion de cada emlpeado del reporte.
' Autor      : FGZ
' Fecha      : 18/08/2005
' Ultima Mod.:            - LA. -  en la consulta buscar todos los paramentros Categoria-puntaje-antig
'                           LA . - Posicion actual en la curva - Antig + Sueldo
'              01-06-2006 - LA. -  tener en cuanta la 4ª coordenada (graduado- no graduado)
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim grilla_val(10) As Single     ' para alojar los valores de:  valgrilla.val(i)
Dim graduado As Integer
Dim Continuar As Boolean
Dim i As Integer
Dim Valor As Double
Dim posActual As String


Dim rs_Stulich As New ADODB.Recordset
Dim rs_cabgrilla As New ADODB.Recordset
Dim rs_valgrilla As New ADODB.Recordset
    
    On Error GoTo M_Error
    
    Flog.writeline ""
    Flog.writeline ""
    Flog.writeline "Ajusto Posicion  y Remuneracion: "
    Flog.writeline "-------------------------------  "
    
    StrSql = " SELECT ternro,empleg, antnum, categ_estrnro,califajus, sueldo, convenio FROM  rep_stulich "
    StrSql = StrSql & " WHERE bpronro = " & NroProceso
    StrSql = StrSql & " ORDER BY empleg ASC "
    If rs_Stulich.State = adStateOpen Then rs_Stulich.Close
    OpenRecordset StrSql, rs_Stulich
    
    Do While Not rs_Stulich.EOF
        Flog.writeline "  "
        Flog.writeline "Legajo " & rs_Stulich!empleg
        
        'Busco en la grilla
        StrSql = "SELECT * FROM cabgrilla "
        StrSql = StrSql & " WHERE cabgrilla.cgrnro = " & NroGrilla
        OpenRecordset StrSql, rs_cabgrilla
        If Not rs_cabgrilla.EOF Then
            If Trim(rs_Stulich!Convenio) = "SI" Then
                graduado = 1
            Else
                graduado = 2
            End If
            
            StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
            StrSql = StrSql & " AND vgrcoor_1 = " & rs_Stulich!Categ_Estrnro
            StrSql = StrSql & " AND vgrcoor_2 = " & rs_Stulich!califajus
            StrSql = StrSql & " AND vgrcoor_3 = " & rs_Stulich!antnum
            StrSql = StrSql & " AND vgrcoor_4 = " & graduado
            '  StrSql = StrSql & " ORDER BY vgrcoor_3 DESC "
            StrSql = StrSql & " ORDER BY vgrorden DESC"
            OpenRecordset StrSql, rs_valgrilla
            If Not rs_valgrilla.EOF Then
                Flog.writeline "Cargo los Valores de la Grilla "
                i = 0
                Do While Not rs_valgrilla.EOF
                    If Not EsNulo(rs_valgrilla!vgrvalor) Then
                        grilla_val(i) = rs_valgrilla!vgrvalor
                        i = i + 1
                    End If
                    rs_valgrilla.MoveNext
                Loop
                
                Flog.writeline "Busco el primer valor no vacio desde el ultimo"
                Valor = 0
                i = 10
                Continuar = True
                Do While i >= 0 And Continuar
                    If grilla_val(i) <> 0 Then
                        Valor = grilla_val(i)
                        Continuar = False
                    End If
                    i = i - 1
                Loop
            Else
                Flog.writeline "No se encontró valor en grilla "
                Flog.writeline "Esta configurado que retorne cero si no lo encuentra "
                Valor = 0
            End If
            rs_valgrilla.Close
                 
            'LA . - Posicion actual en la curva - Antig + Sueldo
            Call DatosPosActualCurva(rs_Stulich!Categ_Estrnro, rs_Stulich!antnum, rs_Stulich!sueldo, graduado, posActual)
            
          
        Else
            Flog.writeline "No se encontró la grilla " & NroGrilla
            Valor = 0
        End If
        
        'Actualizo los valores para poscalif y remupos en el reporte
        StrSql = " UPDATE rep_stulich SET "
        StrSql = StrSql & " posicalif = '" & rs_Stulich!antnum & "." & rs_Stulich!califajus & "'"
        StrSql = StrSql & " , remupos = " & Valor
        StrSql = StrSql & " , poscurva ='" & posActual & "'"
        StrSql = StrSql & " WHERE bpronro =" & NroProceso
        StrSql = StrSql & " AND ternro= " & rs_Stulich!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        rs_Stulich.MoveNext
    Loop
    
Fin:
    If rs_Stulich.State = adStateOpen Then rs_Stulich.Close
    Set rs_Stulich = Nothing
Exit Sub

M_Error:
    Flog.writeline "Error Calculando remuneracion. Error " & Err.Description
    Flog.writeline "Ultimo SQL Ejecutado : " & StrSql
End Sub

 
' _______________________________________________
'  Posicion actual en la curva - Antig + Sueldo
' _______________________________________________
Sub DatosPosActualCurva(Categ_Estrnro, antnum, sueldo, graduado, ByRef posActual)

Dim grilla_val(10) As Single     ' para alojar los valores de:  valgrilla.val(i)
'Dim grilla_coor2(10) As Single   ' para alojar los valores de la coordenada 2 - calific ajustada

'Dim Continuar As Boolean
Dim i As Integer
'Dim Valor As Double

Dim rs_valgrilla As New ADODB.Recordset

Dim anterior  As Double
Dim posterior As Double
Dim promedio As Double

Dim pos1 As Integer
Dim pos2 As Integer

posActual = ""

anterior = 0
posterior = 0
pos1 = 0
pos2 = 0

For i = 1 To 10
    grilla_val(i) = 0
Next

i = 1
    
    StrSql = "SELECT * FROM valgrilla WHERE cgrnro =" & NroGrilla
    StrSql = StrSql & " AND vgrcoor_1 =" & Categ_Estrnro
        'StrSql = StrSql & " AND vgrcoor_2 = " & rs_Stulich!califajus
    StrSql = StrSql & " AND vgrcoor_3 =" & antnum
    StrSql = StrSql & " AND vgrcoor_4 =" & graduado
    StrSql = StrSql & " ORDER BY vgrcoor_2 "
    OpenRecordset StrSql, rs_valgrilla
    If Not rs_valgrilla.EOF Then
        Flog.writeline " Busco la posicion actual en la curva " 'XXXXXXXXXXX
        Do While Not rs_valgrilla.EOF
            If Not EsNulo(rs_valgrilla!vgrcoor_2) Then
                i = rs_valgrilla!vgrcoor_2
                If i < 10 Then
                    If Not EsNulo(rs_valgrilla!vgrvalor) Then
                        grilla_val(i) = rs_valgrilla!vgrvalor
                        'grilla_coor2(i) = rs_valgrilla!vgrcoor2
                        'i = i + 1
                    End If
                End If
            End If
            rs_valgrilla.MoveNext
        Loop
        
        For i = 1 To 5
            If pos1 = 0 Or pos2 = 0 Then
                If grilla_val(i) <= sueldo Then
                    anterior = grilla_val(i)
                    pos1 = i
                Else
                    posterior = grilla_val(i)
                    pos2 = i
                End If
            End If
        Next
         
         ' ver casos extremos............. pos1=0 o pos2=0
        If pos1 = 0 Or pos2 = 0 Then
            If pos1 = 0 Then
                posActual = antnum & ".1 (-)"
            Else
                posActual = antnum & ".5 (+)"
            End If
        Else
            promedio = (anterior + posterior) / 2
            If sueldo > promedio Then
                posActual = antnum & "." & pos2 & " (-)"
            Else
                posActual = antnum & "." & pos1 & " (+)"
            End If
        End If
        'Flog.writeline "Busco el primer valor no vacio desde el ultimo"
    Else
        'Flog.writeline "No se encontró valor en grilla "
        'Flog.writeline "Esta configurado que retorne cero si no lo encuentra "
        'Valor = 0
    End If

    rs_valgrilla.Close
    
End Sub



' _____________________________________________________________________
'  Mapeo de las Siglas de los tipo de Resultados a los valores dados
'  se hace un Mapeo de los resultados de calif RDP para que sea como de Objetivos
' _____________________________________________________________________
Function valorRDP(letra)

Select Case Trim(letra)
   Case "ESE", "EAE":
          valorRDP = 50
   Case "EE":
          valorRDP = 38
   Case "AE":
          valorRDP = 26
   Case "AAE":
          valorRDP = 14
   Case "NAE":
          valorRDP = 0
   Case "N/A", "N/A.":
          valorRDP = 0
   Case Else:
          valorRDP = 0
End Select

End Function
