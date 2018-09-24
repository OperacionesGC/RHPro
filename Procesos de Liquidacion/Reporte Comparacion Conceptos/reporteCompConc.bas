Attribute VB_Name = "repCompConc"
Option Explicit

Global Const Version = "1.01"
Global Const FechaModificacion = "21/10/2009"
Global Const UltimaModificacion = "Encriptacion de string connection"
Global Const UltimaModificacion1 = "Manuel Lopez"

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

Global pliqnro1 As Integer
Global pliqdesc1 As String
Global pliqmesanio1 As String
Global pronro1 As String
Global pliqnro2 As Integer
Global pliqdesc2 As String
Global pliqmesanio2 As String
Global pronro2 As String
Global listaconcnro As String
Global Tabulador As Long

Global IdUser As String
Global Fecha As Date
Global Hora As String

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
Dim rsPeriodos As New ADODB.Recordset
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim rsConceptos As New ADODB.Recordset
Dim I
Dim totalAcum
Dim cantRegistros
Dim PID As String
Dim parametros As String
Dim ArrParametros
Dim concnro As Integer

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
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ReporteCompConceptos" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "Inicio Proceso de Comparación de Conceptos : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
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
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
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
       
       'Obtengo el primer Período
       pliqnro1 = ArrParametros(0)
       
       'Obtengo la lista de procesos asociados al primer Período
       pronro1 = ArrParametros(1)
       
       'Obtengo el segundo Período
       pliqnro2 = ArrParametros(2)
       'listapronro = arrParametros(0)
       
       'Obtengo la lista de procesos asociados al primer Período
       pronro2 = ArrParametros(3)
       
       'Obtengo la lista de conceptos
       listaconcnro = ArrParametros(4)
       
       'EMPIEZA EL PROCESO
       'Busco el periodo desde
       StrSql = "SELECT * FROM periodo WHERE pliqnro = " & pliqnro1
       OpenRecordset StrSql, objRs
        
       If Not objRs.EOF Then
          pliqdesc1 = objRs!pliqDesc
          pliqmesanio1 = objRs!pliqmes & "/" & objRs!pliqanio
       Else
          Flog.writeline "No se encontro el primer Período."
          Exit Sub
       End If
        
       If objRs.State = adStateOpen Then objRs.Close
       
       'Busco el periodo hasta
       StrSql = "SELECT * FROM periodo WHERE pliqnro = " & pliqnro2
       OpenRecordset StrSql, objRs
        
       If Not objRs.EOF Then
          pliqdesc2 = objRs!pliqDesc
          pliqmesanio2 = objRs!pliqmes & "/" & objRs!pliqanio
       Else
          Flog.writeline "No se encontro el segundo Período."
          Exit Sub
       End If
        
       If objRs.State = adStateOpen Then objRs.Close
       
       ' Inserto el encabezado del reporte
       Flog.writeline "Inserto la cabecera del reporte."
       
       StrSql = "INSERT INTO rep_comp_conc (bpronro,pliqnro1,pliqdesc1,pliqmesanio1," & _
                "pronro1,pliqnro2,pliqdesc2,pliqmesanio2,pronro2,iduser,fecha,hora) " & _
                "VALUES (" & _
                 NroProceso & "," & pliqnro1 & ",'" & pliqdesc1 & "','" & pliqmesanio1 & "','" & _
                 pronro1 & "'," & pliqnro2 & ",'" & pliqdesc2 & "','" & pliqmesanio2 & "','" & pronro2 & "'" & ",'" & IdUser & "'," & ConvFecha(Fecha) & "," & "'" & Hora & "')"
       objConn.Execute StrSql, , adExecuteNoRecords
       
       ' Proceso que genera los datos
       Call GenerarDatosConceptoProceso(pronro1, pronro2, listaconcnro)
       
    Else
        Exit Sub
    End If
   
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.writeline "Proceso Finalizado Correctamente"
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
       Flog.writeline "Proceso Incompleto"
    End If
    
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
Sub GenerarDatosConceptoProceso(ByVal pronro1 As String, ByVal pronro2 As String, ByVal lista As String)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset


'Variables donde se guardan los datos del INSERT final
Dim concmonto1 As Double
Dim conccant1 As Double
Dim concmonto2 As Double
Dim conccant2 As Double
Dim difmontoconc As Double
Dim porcmontoconc As Double
Dim difcantconc As Double
Dim porccantconc As Double
Dim concnro As Long
Dim Conccod As String
Dim concabr As String
Dim TConcepto As Integer
Dim Cantidad As Integer
Dim cantidadProcesada As Integer
Dim Orden As Long


On Error GoTo MError

'------------------------------------------------------------------
' Busco los conceptos en la liquidacion
'------------------------------------------------------------------
StrSql = "SELECT proceso.pliqnro, detliq.concnro, SUM(detliq.dlimonto) AS sumdlimonto, "
StrSql = StrSql & " SUM(detliq.dlicant) AS sumdlicant, conccod, concabr,con_acum.acunro "
StrSql = StrSql & " FROM proceso "
StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
StrSql = StrSql & " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro "
StrSql = StrSql & " INNER JOIN concepto ON detliq.concnro = concepto.concnro "
'If listaconcnro = "0" Then
StrSql = StrSql & " INNER JOIN con_acum ON con_acum.concnro = detliq.concnro "
'End If
StrSql = StrSql & " WHERE cabliq.pronro IN (" & pronro1 & "," & pronro2 & ") "
'StrSql = StrSql & " AND concimp = -1 "
If listaconcnro <> "0" Then
    StrSql = StrSql & " AND detliq.concnro IN (" & lista & ") "
Else
    StrSql = StrSql & " AND con_acum.acunro IN (85,86,87,88,89,90,91,92,93,97) "
End If
StrSql = StrSql & "GROUP BY proceso.pliqnro,con_acum.acunro, detliq.concnro, conccod, concabr "
'StrSql = StrSql & "ORDER BY detliq.concnro"
'StrSql = StrSql & "ORDER BY conccod, con_acum.acunro"
StrSql = StrSql & "ORDER BY con_acum.acunro,conccod"
OpenRecordset StrSql, rsConsult

'Seteo el progreso
If rsConsult.RecordCount <> 0 Then
    Cantidad = rsConsult.RecordCount
Else
    Cantidad = 1
End If
IncPorc = 99 / Cantidad
cantidadProcesada = Cantidad

concmonto1 = CDbl(0)
concmonto2 = CDbl(0)
conccant1 = CDbl(0)
conccant2 = CDbl(0)
If Not rsConsult.EOF Then
    concnro = rsConsult!concnro
End If
Orden = 0
Do Until rsConsult.EOF
    
    If rsConsult!concnro <> concnro Then
        Orden = Orden + 1
        ' Genero el procesamiento del Concepto
        Flog.writeline Espacios(Tabulador * 1) & "Insertando en la tabla los datos del concepto " & concnro
        
        TConcepto = CalcularMapeo(Conccod, "CONCEPTOS", 0)
                
        '------------------------------------------------------------------
        ' Realizo los calculos sobre los campos
        '------------------------------------------------------------------
        difmontoconc = CDbl(concmonto1) - CDbl(concmonto2)
        If CDbl(concmonto2) <> 0 Then
            porcmontoconc = (CDbl(difmontoconc) * CDbl(100)) / CDbl(concmonto2)
        Else
            porcmontoconc = IIf(CDbl(concmonto1) <> 0, CDbl(100), CDbl(0))
        End If
        difcantconc = CDbl(conccant1) - CDbl(conccant2)
        If CDbl(conccant2) <> 0 Then
            porccantconc = (CDbl(difcantconc) * CDbl(100)) / CDbl(conccant2)
        Else
            porccantconc = IIf(CDbl(conccant1) <> 0, CDbl(100), CDbl(0))
        End If

        '-------------------------------------------------------------------------------
        'Inserto los datos en la BD
        '-------------------------------------------------------------------------------
        StrSql = "INSERT INTO rep_comp_conc_det (bpronro,tconcepto,concnro,conccod,concabr," & _
                 "concmonto1,conccant1,concmonto2,conccant2,difmontoconc,porcmontoconc," & _
                 "difcantconc,porccantconc,orden) VALUES (" & NroProceso & "," & TConcepto & "," & _
                 concnro & ",'" & Conccod & "','" & concabr & "'," & concmonto1 & "," & _
                 conccant1 & "," & concmonto2 & "," & conccant2 & "," & difmontoconc & "," & _
                 porcmontoconc & "," & difcantconc & "," & porccantconc & "," & Orden & _
                 ")"
        objConn.Execute StrSql, , adExecuteNoRecords

        concmonto1 = CDbl(0)
        concmonto2 = CDbl(0)
        conccant1 = CDbl(0)
        conccant2 = CDbl(0)
    End If
    
    If rsConsult!pliqnro = pliqnro1 Then
        If Not IsNull(rsConsult!sumdlimonto) Then
            concmonto1 = CDbl(rsConsult!sumdlimonto)
        End If
        If Not IsNull(rsConsult!sumdlicant) Then
            conccant1 = CDbl(rsConsult!sumdlicant)
        End If
    End If

    If rsConsult!pliqnro = pliqnro2 Then
        If Not IsNull(rsConsult!sumdlimonto) Then
            concmonto2 = CDbl(rsConsult!sumdlimonto)
        End If
        If Not IsNull(rsConsult!sumdlicant) Then
            conccant2 = CDbl(rsConsult!sumdlicant)
        End If
    End If
    
    concnro = rsConsult!concnro
    Conccod = rsConsult!Conccod
    concabr = rsConsult!concabr
    
    rsConsult.MoveNext
    
    'Actualizo el progreso
    TiempoAcumulado = GetTickCount
    Progreso = Progreso + IncPorc
    cantidadProcesada = cantidadProcesada - 1
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
    StrSql = StrSql & ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
Loop

If rsConsult.State = adStateOpen Then rsConsult.Close

Exit Sub
            
MError:
    Flog.writeline " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub
'----------------------------------------------------------------
'Busca cual es el mapeo de un codigo RHPro a un codigo SAP
'----------------------------------------------------------------
Public Function CalcularMapeo(ByVal Parametro, ByVal Tabla, ByVal Default)

    Dim StrSql As String
    Dim rs_Consult As New ADODB.Recordset
    Dim correcto As Boolean
    Dim Salida
    
    If IsNull(Parametro) Then
       correcto = False
    Else
       correcto = Parametro <> ""
    End If
           
    Salida = Default

    If correcto Then
        StrSql = " SELECT * FROM infotipos_mapeo " & _
                 " WHERE tablaref = '" & Tabla & "' " & _
                 "   AND codinterno = '" & Parametro & "' "
        OpenRecordset StrSql, rs_Consult
        
        If Not rs_Consult.EOF Then
            Salida = CStr(IIf(Not IsNull(rs_Consult!codexterno), rs_Consult!codexterno, Default))
        Else
            Flog.writeline Espacios(Tabulador * 2) & "No se encontró el mapeo para la tabla " & Tabla & " con el codigo interno " & Parametro
        End If
        
        rs_Consult.Close
    Else
        Flog.writeline Espacios(Tabulador * 2) & "Parametro incorrecto al calcular el mapero de la tabla " & Tabla
    End If
    
    CalcularMapeo = Salida

End Function



