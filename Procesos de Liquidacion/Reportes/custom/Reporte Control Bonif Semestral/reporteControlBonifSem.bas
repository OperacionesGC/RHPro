Attribute VB_Name = "repControlBonifSemestral"
Option Explicit

Dim fs, f
Global Flog

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

Global Tabulador As Long

Private Sub Main()

Dim pliqnro As Integer
Dim pliqdesc As String
Dim pliqmesanio As String
Dim pliqdesde As Date
Dim pliqhasta As Date
Dim pliqmes As Integer
Dim pliqanio As Integer
Dim pronro As String
Dim prodesc As String
Dim tprocnro As Integer
Dim IdUser As String
Dim Fecha As Date
Dim Hora As String

Dim NombreArchivo As String
Dim Directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim objRs As New ADODB.Recordset
Dim rsPeriodos As New ADODB.Recordset
Dim param
Dim rsConceptos As New ADODB.Recordset
Dim cantRegistros
Dim PID As String
Dim parametros As String
Dim ArrParametros

    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
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
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    TiempoInicialProceso = GetTickCount
    OpenConnection strconexion, objConn
    
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ReporteControlBonifSemestral" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.Writeline "Inicio Proceso de Control Bonif. Semestral: " & Now
    Flog.Writeline "Cambio el estado del proceso a Procesando"
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.Writeline "PID = " & PID
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.Writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       IdUser = objRs!IdUser
       Fecha = objRs!bprcfecha
       Hora = objRs!bprchora
               
       'Obtengo los parametros del proceso
       parametros = objRs!bprcparam
       ArrParametros = Split(parametros, "@")
       
       'Obtengo el Período
       pliqnro = ArrParametros(0)
       
       ' Obtengo el Tipo de Proceso
       tprocnro = ArrParametros(1)
       
       'Obtengo el Proceso
       pronro = ArrParametros(2)
       
       'EMPIEZA EL PROCESO
       'Busco el periodo
       StrSql = "SELECT * FROM periodo WHERE pliqnro = " & pliqnro
       OpenRecordset StrSql, objRs
        
       If Not objRs.EOF Then
          pliqdesc = objRs!pliqdesc
          pliqmes = objRs!pliqmes
          pliqanio = objRs!pliqanio
          pliqmesanio = pliqmes & "/" & pliqanio
          pliqhasta = objRs!pliqhasta
          pliqdesde = objRs!pliqdesde
       Else
          Flog.Writeline "No se encontro el Período."
          Exit Sub
       End If
        
       If objRs.State = adStateOpen Then objRs.Close
       
       'Busco el proceso
       If CInt(pronro) <> 0 Then
            StrSql = "SELECT * FROM proceso WHERE pronro = " & pronro
            OpenRecordset StrSql, objRs
             
            If Not objRs.EOF Then
               prodesc = objRs!prodesc
            Else
               Flog.Writeline "No se encontro el Proceso."
               Exit Sub
            End If
             
            If objRs.State = adStateOpen Then objRs.Close
       Else
            prodesc = "Todos los Procesos"
       End If
       
       ' Inserto el encabezado del reporte
       Flog.Writeline "Inserto la cabecera del reporte."
       
       StrSql = "INSERT INTO rep_bonif_semestral (bpronro,pliqnro,pliqdesc,pliqmesanio," & _
                "pronro,prodesc,iduser,fecha,hora) " & _
                "VALUES (" & _
                 NroProceso & "," & pliqnro & ",'" & pliqdesc & "','" & pliqmesanio & "'," & _
                 pronro & ",'" & prodesc & "','" & IdUser & "'," & ConvFecha(Fecha) & "," & "'" & Hora & "')"
       objConn.Execute StrSql, , adExecuteNoRecords
       
       ' Proceso que genera los datos
       Call GenerarDatosProceso(tprocnro, pliqnro, pliqmes, pliqanio, pliqdesde, pliqhasta, pronro)
       
    Else
        Exit Sub
    End If
   
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.Writeline "Proceso Finalizado Correctamente"
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
       Flog.Writeline "Proceso Incompleto"
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.Writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.Writeline " Error: " & Err.Description & Now

End Sub
'--------------------------------------------------------------------
' Se encarga de generar los datos
'--------------------------------------------------------------------
Sub GenerarDatosProceso(ByVal tprocnro As Integer, ByVal pliqnro As Integer, ByVal pliqmes As Integer, ByVal pliqanio As Integer, ByVal pliqdesde As Date, ByVal pliqhasta As Date, ByVal pronro As Integer)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset

Dim result_col(14) As Double
Dim col(14, 2) As Long
Dim tenro As String
Dim estrdabr As String
Dim Cantidad As Integer
Dim cantidadProcesada As Integer
Dim i As Integer
Dim lista_conc As String
Dim lista_acum As String
Dim empleg As Long
Dim terape As String
Dim ternom As String
Dim terape2 As String
Dim ternom2 As String
Dim fechaing As Date
Dim catdes As String
Dim catcod As String
Dim j As Integer
Dim total As Double
Dim resultado_1 As Double
Dim resultado_2 As Double
Dim pliqmes_aux As Integer
Dim pliqanio_aux As Integer


On Error GoTo MError

'------------------------------------------------------------------
'Busco los valores del confrep
'------------------------------------------------------------------
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 147 "

OpenRecordset StrSql, rs2

If rs2.EOF Then
   Flog.Writeline "No esta configurado el ConfRep"
   Exit Sub
End If
 
Flog.Writeline "Obtengo los datos del confrep"

i = 1
lista_conc = "0"
lista_acum = "0"
Do Until rs2.EOF

     If rs2!conftipo = "TE" Then
         tenro = rs2!confval
     End If
     
     If rs2!conftipo = "CO" Then
         StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs2!confval
         StrSql = StrSql & " OR conccod = '" & rs2!confval2 & "'"
           
         OpenRecordset StrSql, rs3
           
         If rs3.EOF Then
             col(i, 1) = 0
         Else
             col(i, 1) = rs3!concnro
             lista_conc = lista_conc & "," & CStr(rs3!concnro)
         End If
         col(i, 0) = False
         
         rs3.Close
     End If
     
     If rs2!conftipo = "AC" Then
         col(i, 1) = rs2!confval
         col(i, 0) = True
         lista_acum = lista_acum & "," & CStr(rs2!confval)
     End If
     
     i = i + 1
     rs2.MoveNext
     
Loop
rs2.Close
 
'------------------------------------------------------------------
' Busco los empleados
'------------------------------------------------------------------
StrSql = "SELECT DISTINCT cabliq.empleado "
StrSql = StrSql & " FROM proceso "
StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
StrSql = StrSql & " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro "
StrSql = StrSql & " INNER JOIN con_acum ON con_acum.concnro = detliq.concnro "

StrSql = StrSql & " WHERE proceso.pliqnro = " & pliqnro
If pronro <> 0 Then
    StrSql = StrSql & " AND cabliq.pronro = " & pronro
End If
If tprocnro <> 0 Then
    StrSql = StrSql & " AND proceso.tprocnro = " & tprocnro
End If
'StrSql = StrSql & " WHERE cabliq.pronro = " & pronro

If lista_conc <> "0" And lista_acum <> "0" Then
    StrSql = StrSql & " AND (detliq.concnro IN (" & lista_conc & ") OR con_acum.acunro IN (" & lista_acum & ")) "
Else
    If lista_conc <> "0" Then
        StrSql = StrSql & " AND detliq.concnro IN (" & lista_conc & ") "
    End If
    If lista_acum <> "0" Then
        StrSql = StrSql & " AND con_acum.acunro IN (" & lista_acum & ") "
    End If
End If
StrSql = StrSql & "ORDER BY cabliq.empleado"

OpenRecordset StrSql, rsConsult

'------------------------------------------------------------------
'Seteo el progreso
'------------------------------------------------------------------
If rsConsult.RecordCount <> 0 Then
    Cantidad = rsConsult.RecordCount
Else
    Cantidad = 1
End If
IncPorc = 99 / Cantidad
cantidadProcesada = Cantidad

For i = 2 To 8
    result_col(i) = 0
Next

Do Until rsConsult.EOF
    j = 1
    
    '------------------------------------------------------------------
    ' Busco los datos del Empleado
    '------------------------------------------------------------------
    StrSql = "SELECT empleg, terape, ternom, terape2, ternom2 " & _
             "FROM empleado " & _
             "WHERE ternro = " & rsConsult!Empleado
    
    OpenRecordset StrSql, rs2
    If Not rs2.EOF Then
        empleg = rs2!empleg
        terape = rs2!terape
        ternom = rs2!ternom
        If Not IsNull(rs2!terape2) Then
            terape2 = rs2!terape2
        End If
        If Not IsNull(rs2!ternom2) Then
            ternom2 = rs2!ternom2
        End If
    Else
        Flog.Writeline "No se encontro los datos del empleado " & rsConsult!Empleado
    End If
    rs2.Close
    
    '------------------------------------------------------------------
    ' Busco la Fecha de Ingreso
    '------------------------------------------------------------------
    StrSql = "SELECT altfec " & _
             "FROM fases WHERE fasrecofec = -1 AND empleado = " & rsConsult!Empleado
    
    OpenRecordset StrSql, rs2
    If Not rs2.EOF Then
        fechaing = rs2!altfec
    Else
        rs2.Close
        StrSql = "SELECT altfec " & _
                 "FROM fases WHERE real = -1 AND ternro = " & rsConsult!Empleado
        
        OpenRecordset StrSql, rs2
        If Not rs2.EOF Then
            fechaing = rs2!altfec
        Else
            Flog.Writeline "No se encontro la fecha de ingreso del empleado " & rsConsult!Empleado
            fechaing = Empty
        End If
        
    End If
    rs2.Close
    
    '------------------------------------------------------------------
    ' Busco la estructura
    '------------------------------------------------------------------
    StrSql = "SELECT estrdabr " & _
             "FROM estructura INNER JOIN his_estructura ON estructura.estrnro = his_estructura.estrnro " & _
             "WHERE his_estructura.tenro = " & tenro & " AND his_estructura.ternro = " & rsConsult!Empleado & " AND his_estructura.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his_estructura.htethasta IS NULL OR his_estructura.htethasta >= " & ConvFecha(pliqhasta) & ") "
    
    OpenRecordset StrSql, rs2
    If Not rs2.EOF Then
        estrdabr = rs2!estrdabr
    Else
        Flog.Writeline "No se encontro la estructura del empleado " & rsConsult!Empleado
    End If
    rs2.Close
    
    '------------------------------------------------------------------
    ' Busco la categoria
    '------------------------------------------------------------------
    StrSql = "SELECT estrdabr, estrdext " & _
             "FROM estructura INNER JOIN his_estructura ON estructura.estrnro = his_estructura.estrnro " & _
             "WHERE his_estructura.tenro = 3 AND his_estructura.ternro = " & rsConsult!Empleado & " AND his_estructura.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his_estructura.htethasta IS NULL OR his_estructura.htethasta >= " & ConvFecha(pliqhasta) & ") "
    
    OpenRecordset StrSql, rs2
    If Not rs2.EOF Then
        catdes = IIf(IsNull(rs2!estrdext), "", rs2!estrdext)
        catcod = IIf(IsNull(rs2!estrdabr), "", rs2!estrdabr)
    Else
        Flog.Writeline "No se encontro la categoría del empleado " & rsConsult!Empleado
    End If
    rs2.Close
    
    '-------------------------------------------------------------------------------
    'Busco el valor del concepto para la segunda columna (confnrocol = 2), seis meses para atras
    '-------------------------------------------------------------------------------
    pliqmes_aux = pliqmes
    pliqanio_aux = pliqanio
    result_col(3) = pliqmes_aux
    j = j + 6
    total = 0
    For i = 2 To 7
        pliqmes_aux = pliqmes_aux - 1
        If pliqmes_aux = 0 Then
            pliqmes_aux = 12
            pliqanio_aux = pliqanio_aux - 1
        End If
        
        result_col(j) = 0
        
        If col(2, 0) Then
            StrSql = " SELECT sum(acu_liq.almonto) AS almonto "
            StrSql = StrSql & " FROM acu_liq"
            StrSql = StrSql & " INNER JOIN cabliq ON cabliq.cliqnro = acu_liq.cliqnro "
            StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro "
            StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro AND periodo.pliqmes = " & pliqmes_aux & " AND periodo.pliqanio = " & pliqanio_aux
            StrSql = StrSql & " WHERE acunro = " & col(2, 1)
            StrSql = StrSql & " AND cabliq.empleado = " & rsConsult!Empleado
            StrSql = StrSql & " GROUP BY cabliq.empleado "
        Else
            StrSql = " SELECT sum(detliq.dlimonto) AS almonto "
            StrSql = StrSql & " FROM cabliq "
            StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro "
            StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro AND periodo.pliqmes = " & pliqmes_aux & " AND periodo.pliqanio = " & pliqanio_aux
            StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro "
            StrSql = StrSql & " WHERE cabliq.empleado = " & rsConsult!Empleado & " AND detliq.concnro = " & col(2, 1)
            StrSql = StrSql & " GROUP BY cabliq.empleado "
        End If
        
        OpenRecordset StrSql, rs2
        If Not rs2.EOF Then
            If Not IsNull(rs2!almonto) Then
                result_col(j) = rs2!almonto
                total = total + rs2!almonto
            End If
        Else
            Flog.Writeline "No se encontro valor para el empleado " & rsConsult!Empleado & " en el periodo " & pliqmes & "/" & pliqanio & " para la 2 columna confrep"
        End If
        rs2.Close
    
        j = j - 1
        
    Next
    j = j + 6
    
    '------------------------------------------------------------------
    ' Calculo el promedio
    '------------------------------------------------------------------
    j = j + 1
    result_col(j) = total / 6
    
    '------------------------------------------------------------------
    ' Busco el valor de la 3 y 4 columna
    '------------------------------------------------------------------
    For i = 3 To 4
        If col(i, 0) Then
            StrSql = " SELECT sum(almonto) "
            StrSql = StrSql & " FROM acu_liq"
            StrSql = StrSql & " INNER JOIN cabliq   ON acu_liq.cliqnro = cabliq.cliqnro "
            StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro "
            If pronro <> 0 Then
                StrSql = StrSql & "AND proceso.pronro = " & pronro
            End If
            If tprocnro <> 0 Then
                StrSql = StrSql & " AND proceso.tprocnro = " & tprocnro
            End If
            StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro AND periodo.pliqnro = " & pliqnro
            StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & rsConsult!Empleado
            StrSql = StrSql & " WHERE acu_liq.acunro = " & col(i, 1)
            StrSql = StrSql & " GROUP BY cabliq.empleado "
        Else
            StrSql = " SELECT sum(detliq.dlimonto) AS almonto "
            StrSql = StrSql & " FROM cabliq "
            StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro "
            If pronro <> 0 Then
                StrSql = StrSql & "AND proceso.pronro = " & pronro
            End If
            If tprocnro <> 0 Then
                StrSql = StrSql & " AND proceso.tprocnro = " & tprocnro
            End If
            StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro AND periodo.pliqnro = " & pliqnro
            StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & rsConsult!Empleado & " AND detliq.concnro = " & col(i, 1)
            StrSql = StrSql & " GROUP BY cabliq.empleado "
        End If
        
        j = j + 1
        
        OpenRecordset StrSql, rs2
        If Not rs2.EOF Then
            result_col(j) = rs2!almonto
        Else
            Flog.Writeline "No se encontro valor para el empleado " & rsConsult!Empleado & " para la " & i & " columna confrep"
        End If
        rs2.Close
    Next
    
    '-------------------------------------------------------------------------------
    'Busco si el empleado tiene el mismo valor para los conceptos definidos en la 5 y 6 columna
    '-------------------------------------------------------------------------------
    j = j + 1
    result_col(j) = 1
    
    If col(5, 0) Then
        StrSql = " SELECT sum(almonto) "
        StrSql = StrSql & " FROM acu_liq"
        StrSql = StrSql & " INNER JOIN cabliq   ON acu_liq.cliqnro = cabliq.cliqnro "
        StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro "
        If pronro <> 0 Then
            StrSql = StrSql & "AND proceso.pronro = " & pronro
        End If
        If tprocnro <> 0 Then
            StrSql = StrSql & " AND proceso.tprocnro = " & tprocnro
        End If
        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro AND periodo.pliqnro = " & pliqnro
        StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & rsConsult!Empleado
        StrSql = StrSql & " WHERE acu_liq.acunro = " & col(5, 1)
        StrSql = StrSql & " GROUP BY cabliq.empleado "
    Else
        StrSql = " SELECT sum(detliq.dlimonto) AS almonto "
        StrSql = StrSql & " FROM cabliq "
        StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro "
        If pronro <> 0 Then
            StrSql = StrSql & "AND proceso.pronro = " & pronro
        End If
        If tprocnro <> 0 Then
            StrSql = StrSql & " AND proceso.tprocnro = " & tprocnro
        End If
        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro AND periodo.pliqnro = " & pliqnro
        StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & rsConsult!Empleado & " AND detliq.concnro = " & col(5, 1)
        StrSql = StrSql & " GROUP BY cabliq.empleado "
    End If
    
    OpenRecordset StrSql, rs2
    If Not rs2.EOF Then
        resultado_1 = rs2!almonto
    Else
        Flog.Writeline "No se encontro valor para el empleado " & rsConsult!Empleado & " para la 5 columna confrep"
    End If
    rs2.Close
    
    If col(6, 0) Then
        StrSql = " SELECT sum(almonto) "
        StrSql = StrSql & " FROM acu_liq"
        StrSql = StrSql & " INNER JOIN cabliq   ON acu_liq.cliqnro = cabliq.cliqnro "
        StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro "
        If pronro <> 0 Then
            StrSql = StrSql & "AND proceso.pronro = " & pronro
        End If
        If tprocnro <> 0 Then
            StrSql = StrSql & " AND proceso.tprocnro = " & tprocnro
        End If
        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro AND periodo.pliqnro = " & pliqnro
        StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & rsConsult!Empleado
        StrSql = StrSql & " WHERE acu_liq.acunro = " & col(6, 1)
        StrSql = StrSql & " GROUP BY cabliq.empleado "
    Else
        StrSql = " SELECT sum(detliq.dlimonto) AS almonto "
        StrSql = StrSql & " FROM cabliq "
        StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro "
        If pronro <> 0 Then
            StrSql = StrSql & "AND proceso.pronro = " & pronro
        End If
        If tprocnro <> 0 Then
            StrSql = StrSql & " AND proceso.tprocnro = " & tprocnro
        End If
        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro AND periodo.pliqnro = " & pliqnro
        StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & rsConsult!Empleado & " AND detliq.concnro = " & col(6, 1)
        StrSql = StrSql & " GROUP BY cabliq.empleado "
    End If
    
    OpenRecordset StrSql, rs2
    If Not rs2.EOF Then
        resultado_2 = rs2!almonto
    Else
        Flog.Writeline "No se encontro valor para el empleado " & rsConsult!Empleado & " para la 6 columna confrep"
    End If
    rs2.Close
    
    If CDbl(resultado_1) = CDbl(resultado_2) Then
        result_col(j) = 0
    End If
    
    '-------------------------------------------------------------------------------
    'Busco la 7 columna del confrep
    '-------------------------------------------------------------------------------
    j = j + 1
    result_col(j) = 0
    
    If col(7, 0) Then
        StrSql = " SELECT sum(almonto) "
        StrSql = StrSql & " FROM acu_liq"
        StrSql = StrSql & " INNER JOIN cabliq   ON acu_liq.cliqnro = cabliq.cliqnro "
        StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro "
        If pronro <> 0 Then
            StrSql = StrSql & "AND proceso.pronro = " & pronro
        End If
        If tprocnro <> 0 Then
            StrSql = StrSql & " AND proceso.tprocnro = " & tprocnro
        End If
        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro AND periodo.pliqnro = " & pliqnro
        StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & rsConsult!Empleado
        StrSql = StrSql & " WHERE acu_liq.acunro = " & col(7, 1)
        StrSql = StrSql & " GROUP BY cabliq.empleado "
    Else
        StrSql = " SELECT sum(detliq.dlimonto) AS almonto "
        StrSql = StrSql & " FROM cabliq "
        StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro "
        If pronro <> 0 Then
            StrSql = StrSql & "AND proceso.pronro = " & pronro
        End If
        If tprocnro <> 0 Then
            StrSql = StrSql & " AND proceso.tprocnro = " & tprocnro
        End If
        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro AND periodo.pliqnro = " & pliqnro
        StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & rsConsult!Empleado & " AND detliq.concnro = " & col(7, 1)
        StrSql = StrSql & " GROUP BY cabliq.empleado "
    End If
    
    OpenRecordset StrSql, rs2
    If Not rs2.EOF Then
        result_col(j) = rs2!almonto / 2
    Else
        Flog.Writeline "No se encontro valor para el empleado " & rsConsult!Empleado & " para la 7 columna confrep"
    End If
    rs2.Close
    
    '-------------------------------------------------------------------------------
    'Busco la 8 columna del confrep
    '-------------------------------------------------------------------------------
    j = j + 1
    result_col(j) = 0
    
    If col(8, 0) Then
        StrSql = " SELECT sum(almonto) "
        StrSql = StrSql & " FROM acu_liq"
        StrSql = StrSql & " INNER JOIN cabliq   ON acu_liq.cliqnro = cabliq.cliqnro "
        StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro "
        If pronro <> 0 Then
            StrSql = StrSql & "AND proceso.pronro = " & pronro
        End If
        If tprocnro <> 0 Then
            StrSql = StrSql & " AND proceso.tprocnro = " & tprocnro
        End If
        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro AND periodo.pliqnro = " & pliqnro
        StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & rsConsult!Empleado
        StrSql = StrSql & " WHERE acu_liq.acunro = " & col(8, 1)
        StrSql = StrSql & " GROUP BY cabliq.empleado "
    Else
        StrSql = " SELECT sum(detliq.dlimonto) AS almonto "
        StrSql = StrSql & " FROM cabliq "
        StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro "
        If pronro <> 0 Then
            StrSql = StrSql & "AND proceso.pronro = " & pronro
        End If
        If tprocnro <> 0 Then
            StrSql = StrSql & " AND proceso.tprocnro = " & tprocnro
        End If
        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro AND periodo.pliqnro = " & pliqnro
        StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & rsConsult!Empleado & " AND detliq.concnro = " & col(8, 1)
        StrSql = StrSql & " GROUP BY cabliq.empleado "
    End If
    
    OpenRecordset StrSql, rs2
    If Not rs2.EOF Then
        result_col(j) = rs2!almonto
    Else
        Flog.Writeline "No se encontro valor para el empleado " & rsConsult!Empleado & " para la 8 columna confrep"
    End If
    rs2.Close
    
    
    '-------------------------------------------------------------------------------
    'Inserto los datos en la BD
    '-------------------------------------------------------------------------------
    StrSql = "INSERT INTO rep_bonif_semestral_det (bpronro,ternro,empleg,terape,ternom,terape2," & _
             "ternom2,fechaing,col1,col2,col3,col4,col5,col6,col7,col8,col9,col10," & _
             "col11,col12,col13,col14,catdes,catcod)" & _
             " VALUES (" & NroProceso & "," & rsConsult!Empleado & "," & empleg & ",'" & _
             terape & "','" & ternom & "','" & terape2 & "','" & ternom2 & "'," & _
             ConvFecha(fechaing) & ",'" & estrdabr & "'," & numberForSQL(result_col(2)) & "," & numberForSQL(result_col(3)) & "," & _
             numberForSQL(result_col(4)) & "," & numberForSQL(result_col(5)) & "," & numberForSQL(result_col(6)) & "," & _
             numberForSQL(result_col(7)) & "," & numberForSQL(result_col(8)) & "," & _
             numberForSQL(result_col(9)) & "," & numberForSQL(result_col(10)) & "," & _
             numberForSQL(result_col(11)) & "," & numberForSQL(result_col(12)) & "," & _
             numberForSQL(result_col(13)) & "," & numberForSQL(result_col(14)) & ",'" & _
             catdes & "','" & catcod & "')"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    terape = ""
    ternom = ""
    terape2 = ""
    ternom2 = ""
    estrdabr = ""
    For i = 2 To 14
        result_col(i) = 0
    Next
    
    rsConsult.MoveNext
    
    'Actualizo el progreso
    TiempoAcumulado = GetTickCount
    Progreso = Progreso + IncPorc
    cantidadProcesada = cantidadProcesada - 1
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & numberForSQL(Progreso)
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
    StrSql = StrSql & ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
Loop

If rsConsult.State = adStateOpen Then rsConsult.Close
If rs2.State = adStateOpen Then rs2.Close
If rs3.State = adStateOpen Then rs3.Close

Exit Sub
            
MError:
    Flog.Writeline " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub

Function numberForSQL(Str)
   
  numberForSQL = Replace(Str, ",", ".")

End Function

