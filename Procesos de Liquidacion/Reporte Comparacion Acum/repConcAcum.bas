Attribute VB_Name = "repCompAcum"
Option Explicit

'Version: 1.00
'

'Const Version = 1.01
'Const FechaVersion = "27/10/2008" ' Se corrigió para que saliera una linea por empleado

'Global Const Version = "1.02" ' Cesar Stankunas
'Global Const FechaVersion = "05/08/2009"
'Global Const UltimaModificacion = ""    'Encriptacion de string connection

'Global Const Version = "1.03" ' Sebastian Stremel
'Global Const FechaVersion = "26/09/2011"
'Global Const UltimaModificacion = "se cambio el tipo de la variable ternro de integer a long porque provocaba desbordamiento"

'Global Const Version = "1.04" ' Sebastian Stremel
'Global Const FechaVersion = "14/09/2012"
'Global Const UltimaModificacion = "correcion reporte por concepto y por empleado - CAS-16664 - Sykes - Actualizacion de Reporte Comparativo - Rechazo"

'Global Const Version = "1.05" ' Lisandro Moro
'Global Const FechaVersion = "12/12/2014"
'Global Const UltimaModificacion = "CAS-28254 - ACARA - Error en Generacion reporte comparativo de liquidacion. - Se cambio el tipo de la variable ternro de integer a long porque provocaba desbordamiento"

'Global Const Version = "1.06" ' Miriam Ruiz
'Global Const FechaVersion = "13/01/2015"
'Global Const UltimaModificacion = "CAS-28254 - ACARA - Error en Generacion reporte comparativo de liquidacion. - Se limpió el log y se corrigió la barra de progreso para que aparezca una sola vez"

Global Const Version = "1.07"
Global Const FechaVersion = "05/05/2016"
Global Const UltimaModificacion = "FMD - CAS-36720 - MONASTERIO BASE 1 - Bug en reporte comparativo - Se corrige carga de conceptos modificando consulta para que no traiga empleados con datos nulos"

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
Global listaacunro As String
Global iguales As Boolean
Global cantRegistros As Long
Global totalAcum As Long





Private Sub Main()

Dim NombreArchivo As String
Dim directorio As String
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
Dim rsAcum As New ADODB.Recordset
Dim I
Dim PID As String
Dim parametros As String
Dim ArrParametros
Dim acuNro As Integer
Dim ContAux As Integer
Dim ContAux2 As Integer


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
    
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ReporteCompAcum" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    
    On Error Resume Next
    OpenConnection strconexion, objConn
    OpenConnection strconexion, objconn2
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    Flog.writeline "Inicio Proceso de Comparación de Acumuladores : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconn2.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       
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
       
       'Obtengo la lista de acumuladores
       listaacunro = ArrParametros(4)
       
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
       
       'Obtengo los acumuladores sobre los que deberia generar la comparación
       Flog.writeline "CargarAcum(pronro1, pronro2, rsacum) = " & pronro1 & " , " & pronro2
       Call CargarAcum(pronro1, pronro2, rsAcum)
       Flog.writeline "PASO POR CargarAcum"
       
       cantRegistros = CLng(rsAcum.RecordCount)
       totalAcum = CLng(cantRegistros)
       ContAux = cantRegistros
       ContAux2 = cantRegistros
       Dim vbPASO As Integer
       ' Genero los datos
       Do Until rsAcum.EOF
            EmpErrores = False
            acuNro = rsAcum!acuNro
              
            ' Genero el procesamiento del acumulador
            Flog.writeline "Generando datos acumulador " & acuNro
                      
            ' Genero los datos del acumulador
            Flog.writeline "Call GenerarDatosAcumuladorProceso(acunro = " & acuNro & ", pronro1 = " & pronro1 & ", pronro2 = " & pronro2 & ")"
            Call GenerarDatosAcumuladorProceso(acuNro, pronro1, pronro2)
            vbPASO = vbPASO + 1
            Flog.writeline "PASO POR GenerarDatosAcumulador " & vbPASO & " veces"
            
            
            TiempoAcumulado = GetTickCount
              
           ' cantRegistros = cantRegistros - 1
            ContAux = ContAux - 1
            
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((ContAux2 - ContAux) * 100) / ContAux2) & _
                     ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                     " WHERE bpronro = " & NroProceso
                 
            Flog.writeline StrSql
            objconn2.Execute StrSql, , adExecuteNoRecords
             
            rsAcum.MoveNext
       Loop
    Else
        Exit Sub
    End If
   
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.writeline "Proceso Finalizado Correctamente"
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =101, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
       Flog.writeline "Proceso Incompleto"
    End If
    
    objconn2.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now

End Sub

'--------------------------------------------------------------------
' Busco la sumatoria de los acu_liq para los procesos del primer y segundo Período
'--------------------------------------------------------------------
Sub BuscarAcu_liq(ByVal acuNro As Integer, ByVal pronro As String, ByRef almonto As Double, ByRef alcant As Double)
Dim StrAcum As String
Dim rsAculiq As New ADODB.Recordset

    StrAcum = "SELECT sum(acu_liq.almonto) totalmonto, sum(acu_liq.alcant) totalcant " & _
              " FROM acu_liq " & _
              " INNER JOIN cabliq ON cabliq.cliqnro = acu_liq.cliqnro " & _
              " WHERE cabliq.pronro IN (" & pronro & ") " & _
              " AND acu_liq.acunro = " & acuNro
    
  '  Flog.writeline "/////////////////////////////////////////////////////////////////////////////////////////////"
  '  Flog.writeline "StrAcum = " & StrAcum
    
    OpenRecordset StrAcum, rsAculiq
    almonto = 0
    alcant = 0
    If Not rsAculiq.EOF Then
        If Not IsNull(rsAculiq!totalmonto) Then
            almonto = CDbl(rsAculiq!totalmonto)
        End If
        If Not IsNull(rsAculiq!totalcant) Then
            alcant = CDbl(rsAculiq!totalcant)
        End If
    End If
    
   ' Flog.writeline "almonto = " & almonto
   ' Flog.writeline "alcant = " & alcant
    
   ' Flog.writeline "/////////////////////////////////////////////////////////////////////////////////////////////"
    If rsAculiq.State = adStateOpen Then rsAculiq.Close

End Sub
'--------------------------------------------------------------------
' Se encarga de generar la comparacion para el acumulador
'--------------------------------------------------------------------
Sub GenerarDatosAcumuladorProceso(ByVal acuNro As Integer, ByVal pronro1 As String, ByVal pronro2 As String)

Dim StrSql As String
Dim rsConceptos As New ADODB.Recordset
Dim rsEmpleados As New ADODB.Recordset
Dim rsConsult As New ADODB.Recordset
Dim rsConsult2 As New ADODB.Recordset


'Variables donde se guardan los datos del INSERT final
Dim acumonto1 As Double
Dim acucant1 As Double
Dim acumonto2 As Double
Dim acucant2 As Double
Dim concmonto1 As Double
Dim conccant1 As Double
Dim concmonto2 As Double
Dim conccant2 As Double
Dim empmonto1 As Double
Dim empcant1 As Double
Dim empmonto2 As Double
Dim empcant2 As Double
Dim difmontoacu As Double
Dim porcmontoacu As Double
Dim difcantacu As Double
Dim porccantacu As Double
Dim difmontoconc As Double
Dim porcmontoconc As Double
Dim difcantconc As Double
Dim porccantconc As Double
Dim difmontoemp As Double
Dim porcmontoemp As Double
Dim difcantemp As Double
Dim porccantemp As Double

Dim ConcNro As Integer
' Dim Ternro As Integer provocaba desbordamiento
Dim Ternro As Long

Dim empleg As Long
Dim terape As String
Dim terape2 As String
Dim ternom As String
Dim ternom2 As String
Dim ConcCod As String
Dim concabr As String
Dim acudesc As String
Dim Empleado As String
Dim Seguir As Boolean

On Error GoTo MError

'------------------------------------------------------------------
'Busco los datos del acumulador
'------------------------------------------------------------------
StrSql = " SELECT acudesabr FROM acumulador WHERE acunro = " & acuNro
            
OpenRecordset StrSql, rsConsult
acudesc = ""
If Not rsConsult.EOF Then
    acudesc = rsConsult!acudesabr
Else
    acudesc = "NULL"
End If
If rsConsult.State = adStateOpen Then rsConsult.Close

'------------------------------------------------------------------
' Busco los valores en acu_liq de los procesos del primer Período
'------------------------------------------------------------------
Flog.writeline "Call BuscarAcu_liq(acunro, pronro1, acumonto1, acucant1) - 1"

Call BuscarAcu_liq(acuNro, pronro1, acumonto1, acucant1)

Flog.writeline "PASO POR BuscarAcu_liq 1"
'------------------------------------------------------------------
' Busco los valores en acu_liq de los procesos del segundo Período
'------------------------------------------------------------------
Flog.writeline "Call BuscarAcu_liq(acunro, pronro2, acumonto2, acucant2) - 2"

'Flog.writeline "********************************************************************"
'Flog.writeline "PARAMETROS DE 2"

'Flog.writeline acuNro
'Flog.writeline pronro2
'Flog.writeline acumonto2
'Flog.writeline acucant2
'Flog.writeline "**********************************************************************"
Call BuscarAcu_liq(acuNro, pronro2, acumonto2, acucant2)

Flog.writeline "PASO POR BuscarAcu_liq 2"
'------------------------------------------------------------------
' Realizo el calculo sobre el acumulador
'------------------------------------------------------------------
difmontoacu = CDbl(acumonto1) - CDbl(acumonto2)

'Flog.writeline "difmontoacu= " & difmontoacu
'Flog.writeline "acumonto1= " & acumonto1
'Flog.writeline "acumonto2= " & acumonto2

If CDbl(acumonto2) <> 0 Then
    porcmontoacu = (CDbl(difmontoacu) * CDbl(100)) / CDbl(acumonto2)
Else
    porcmontoacu = IIf(CDbl(acumonto1) <> 0, CDbl(100), CDbl(0))
End If
difcantacu = CDbl(acucant1) - CDbl(acucant2)
If CDbl(acucant2) <> 0 Then
    porccantacu = (CDbl(difcantacu) * CDbl(100)) / CDbl(acucant2)
Else
    porccantacu = IIf(CDbl(acucant1) <> 0, CDbl(100), CDbl(0))
End If
            
Flog.writeline "porccantacu" & porccantacu
'------------------------------------------------------------------
' Busco los empleados en la liquidacion
'------------------------------------------------------------------
'StrSql = "SELECT proceso.pliqnro, cabliq.empleado, detliq.concnro, " & _
'         "SUM(detliq.dlimonto) sumdlimonto, SUM(detliq.dlicant) sumdlicant, " & _
'         "empleg, terape, terape2, ternom, ternom2, conccod, concabr " & _
'         "From proceso " & _
'         "INNER JOIN cabliq ON proceso.pronro = cabliq.pronro " & _
'         "INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro " & _
'         "INNER JOIN con_acum ON detliq.concnro = con_acum.concnro " & _
'         "LEFT JOIN v_empleado ON cabliq.empleado = v_empleado.ternro " & _
'         "LEFT JOIN concepto ON detliq.concnro = concepto.concnro " & _
'         "WHERE cabliq.pronro IN (" & pronro1 & "," & pronro2 & ") " & _
'         "AND con_acum.acunro = " & acunro & " " & _
'         "GROUP BY proceso.pliqnro, cabliq.empleado, detliq.concnro, " & _
'         "empleg, terape, terape2, ternom, ternom2, conccod, concabr " & _
'         "ORDER BY detliq.concnro, cabliq.empleado"
         
StrSql = "SELECT proceso.pliqnro,proceso.pronro, cabliq.empleado, detliq.concnro, " & _
         "SUM(detliq.dlimonto) sumdlimonto, SUM(detliq.dlicant) sumdlicant, " & _
         "empleg, terape, terape2, ternom, ternom2, conccod, concabr " & _
         " From proceso " & _
         " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro " & _
         " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro " & _
         " INNER JOIN con_acum ON detliq.concnro = con_acum.concnro " & _
         " INNER JOIN empleado ON cabliq.empleado = empleado.ternro " & _
         " INNER JOIN concepto ON detliq.concnro = concepto.concnro " & _
         " WHERE cabliq.pronro IN (" & pronro1 & "," & pronro2 & ") " & _
         " AND con_acum.acunro = " & acuNro & " " & _
         " GROUP BY proceso.pronro,proceso.pliqnro, cabliq.empleado, detliq.concnro, " & _
         " empleg, terape, terape2, ternom, ternom2, conccod, concabr " & _
         " ORDER BY detliq.concnro, cabliq.empleado"
         
         Flog.writeline StrSql
OpenRecordset StrSql, rsConsult

concmonto1 = CDbl(0)
concmonto2 = CDbl(0)
conccant1 = CDbl(0)
conccant2 = CDbl(0)

cantRegistros = CLng(rsConsult.RecordCount)
totalAcum = CLng(cantRegistros)

Do Until rsConsult.EOF

    Empleado = rsConsult!Empleado
    ConcNro = rsConsult!ConcNro
    Ternro = Empleado
    
   ' Flog.writeline "Empleado= " & Empleado
   ' Flog.writeline "concnro= " & ConcNro
   ' Flog.writeline "ternro= " & Ternro
    
  ' ' Flog.writeline "#############################################"
    'Flog.writeline "TOMO EL NOMBRE DEL TERCERO - SI NO TIENE GENERA BRONCA"
   ' Flog.writeline "rsConsult!ternom=" & rsConsult!ternom
    
    ternom = rsConsult!ternom
   ' Flog.writeline "#############################################"
    
   ' Flog.writeline "ternom = rsConsult!ternom == " & ternom
    
    If IsNull(rsConsult!ternom2) Then
        ternom2 = ""
    Else
        ternom2 = rsConsult!ternom2
    End If
    
    'Flog.writeline "ternom2 = " & ternom2
    
    terape = rsConsult!terape
    'Flog.writeline "terape=" & terape
    
    If IsNull(rsConsult!terape2) Then
        terape2 = ""
    Else
        terape2 = rsConsult!terape2
    End If
    'Flog.writeline "terape2 = " & terape2
    If IsNull(rsConsult!empleg) Then
        Flog.writeline "empleg ES NULL para el ternro " & Ternro
        empleg = 0
    Else
        empleg = rsConsult!empleg
    End If
    
    ConcCod = rsConsult!ConcCod
    concabr = rsConsult!concabr
    
    'Flog.writeline "Conccod = " & ConcCod
    'Flog.writeline "concabr = " & concabr
    
    empmonto1 = 0
    empcant1 = 0
    
    'Flog.writeline "pliqnro1 = " & pliqnro1
    iguales = False
    
    'me fijo si periodo 1 y periodo 2 son los mismos
        If pliqnro1 = pliqnro2 Then
            iguales = True
        End If
    'hasta aca
    
    If iguales Then
        If rsConsult!pliqnro = pliqnro1 And rsConsult!pronro = pronro1 Then
            'Flog.writeline "rsConsult!sumdlimonto = " & rsConsult!sumdlimonto
            If Not IsNull(rsConsult!sumdlimonto) Then
                empmonto1 = CDbl(rsConsult!sumdlimonto)
            End If
            
            'Flog.writeline "rsConsult!sumdlicant = " & rsConsult!sumdlicant
            If Not IsNull(rsConsult!sumdlicant) Then
                empcant1 = CDbl(rsConsult!sumdlicant)
            End If
            'Flog.writeline "empmonto1 = " & empmonto1
            'Flog.writeline "concmonto1 = " & concmonto1
            
            concmonto1 = concmonto1 + CDbl(empmonto1)
            
            'Flog.writeline "conccant1 = " & conccant1
            'Flog.writeline "empcant1 = " & empcant1
            
            conccant1 = conccant1 + CDbl(empcant1)
            rsConsult.MoveNext
        End If
    Else
        If rsConsult!pliqnro = pliqnro1 Then
            'Flog.writeline "rsConsult!sumdlimonto = " & rsConsult!sumdlimonto
            If Not IsNull(rsConsult!sumdlimonto) Then
                empmonto1 = CDbl(rsConsult!sumdlimonto)
            End If
            
           ' Flog.writeline "rsConsult!sumdlicant = " & rsConsult!sumdlicant
            If Not IsNull(rsConsult!sumdlicant) Then
                empcant1 = CDbl(rsConsult!sumdlicant)
            End If
            'Flog.writeline "empmonto1 = " & empmonto1
           ' Flog.writeline "concmonto1 = " & concmonto1
            
            concmonto1 = concmonto1 + CDbl(empmonto1)
            
           ' Flog.writeline "conccant1 = " & conccant1
            'Flog.writeline "empcant1 = " & empcant1
            
            conccant1 = conccant1 + CDbl(empcant1)
            rsConsult.MoveNext
        End If
    End If
    empmonto2 = 0
    empcant2 = 0
    If Not rsConsult.EOF Then
        'Flog.writeline "pliqnro2 = " & pliqnro2
        'Flog.writeline "Empleado = " & Empleado
       ' Flog.writeline "concnro = " & ConcNro
        
        If rsConsult!pliqnro = pliqnro2 And rsConsult!Empleado = Empleado And rsConsult!ConcNro = ConcNro Then
            If Not IsNull(rsConsult!sumdlimonto) Then
                empmonto2 = CDbl(rsConsult!sumdlimonto)
            End If
            If Not IsNull(rsConsult!sumdlicant) Then
                empcant2 = CDbl(rsConsult!sumdlicant)
            End If
            'Flog.writeline "empcant2 = " & empcant2
            
            concmonto2 = concmonto2 + CDbl(empmonto2)
           ' Flog.writeline "concmonto2 = " & concmonto2
            conccant2 = conccant2 + CDbl(empcant2)
            'Flog.writeline "conccant2 = " & conccant2
            rsConsult.MoveNext
        End If
    End If
    
    'Flog.writeline "If (empmonto1 <> 0 Or empcant1 <> 0 Or empmonto2 <> 0 Or empcant2 <> 0) Then"
    If (empmonto1 <> 0 Or empcant1 <> 0 Or empmonto2 <> 0 Or empcant2 <> 0) Then
        '------------------------------------------------------------------
        ' Realizo los calculos sobre los campos
        '------------------------------------------------------------------
        difmontoemp = CDbl(empmonto1) - CDbl(empmonto2)
      '  Flog.writeline "difmontoemp = " & difmontoemp
        If CDbl(empmonto2) <> 0 Then
            porcmontoemp = (CDbl(difmontoemp) * CDbl(100)) / CDbl(empmonto2)
        Else
            porcmontoemp = IIf(CDbl(empmonto1) <> 0, CDbl(100), CDbl(0))
        End If
       ' Flog.writeline "porcmontoemp = " & porcmontoemp
        
        difcantemp = CDbl(empcant1) - CDbl(empcant2)
       ' Flog.writeline "difcantemp = " & difcantemp
        If CDbl(empcant2) <> 0 Then
            porccantemp = (CDbl(difcantemp) * CDbl(100)) / CDbl(empcant2)
        Else
            porccantemp = IIf(CDbl(empcant1) <> 0, CDbl(100), CDbl(0))
        End If
        'Flog.writeline "porccantemp = " & porccantemp
        '-------------------------------------------------------------------------------
        'Inserto los datos en la BD
        '-------------------------------------------------------------------------------
        StrSql = "SELECT *"
        StrSql = StrSql & " FROM rep_comp_acum"
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND ternro = " & Ternro
        StrSql = StrSql & " AND acunro = " & acuNro
        StrSql = StrSql & " AND concnro = " & ConcNro
        OpenRecordset StrSql, rsConsult2
        If rsConsult2.EOF Then
        
            'Flog.writeline "Inserto los datos en la BD"
            
            StrSql = "INSERT INTO rep_comp_acum (bpronro,acunro,acudesc,concnro,conccod," & _
                     "concabr,ternro,empleg,terape,ternom2,terape2,ternom,pliqnro1," & _
                     "pliqdesc1,pliqmesanio1,acumonto1,acucant1,concmonto1,conccant1," & _
                     "empmonto1,empcant1,pliqnro2,pliqdesc2,pliqmesanio2,acumonto2," & _
                     "acucant2,concmonto2,conccant2,empmonto2,empcant2,difmontoacu," & _
                     "porcmontoacu,difcantacu,porccantacu,difmontoconc,porcmontoconc," & _
                     "difcantconc,porccantconc,difmontoemp,porcmontoemp,difcantemp," & _
                     "porccantemp,pronro1,pronro2) "
            StrSql = StrSql & "VALUES (" & _
                     NroProceso & "," & acuNro & ",'" & acudesc & "'," & ConcNro & ",'" & _
                     ConcCod & "','" & concabr & "'," & Ternro & ",'" & empleg & "','" & _
                     terape & "','" & ternom2 & "','" & terape2 & "','" & ternom & "'," & _
                     pliqnro1 & ",'" & pliqdesc1 & "','" & pliqmesanio1 & "'," & acumonto1 & "," & _
                     acucant1 & ",0,0," & empmonto1 & "," & _
                     empcant1 & "," & pliqnro2 & ",'" & pliqdesc2 & "','" & pliqmesanio2 & "'," & _
                     acumonto2 & "," & acucant2 & ",0,0," & _
                     empmonto2 & "," & empcant2 & "," & difmontoacu & "," & porcmontoacu & "," & _
                     difcantacu & "," & porccantacu & ",0,0,0,0," & _
                     difmontoemp & "," & porcmontoemp & "," & _
                     difcantemp & "," & porccantemp & ",'" & pronro1 & "','" & pronro2 & "'" & _
                     ")"
           ' Flog.writeline StrSql
            objConn.Execute StrSql, , adExecuteNoRecords
    Else
    
        'Flog.writeline "Modifico los datos en la BD"
        
        difmontoemp = (rsConsult2!empmonto1 + empmonto1) - (rsConsult2!empmonto2 + empmonto2)
        difcantemp = (rsConsult2!empcant1 + empcant1) - (rsConsult2!empcant2 + empcant2)
        If CDbl((rsConsult2!empmonto2 + empmonto2)) <> 0 Then
            porcmontoemp = (CDbl(difmontoemp) * CDbl(100)) / CDbl((rsConsult2!empmonto2 + empmonto2))
        Else
            porcmontoemp = IIf(CDbl((rsConsult2!empmonto1 + empmonto1)) <> 0, CDbl(100), CDbl(0))
        End If
        Flog.writeline "porcmontoemp = " & porcmontoemp
        If CDbl((rsConsult2!empcant2 + empcant2)) <> 0 Then
            porccantemp = (CDbl(difcantemp) * CDbl(100)) / CDbl((rsConsult2!empcant2 + empcant2))
        Else
            porccantemp = IIf(CDbl((rsConsult2!empcant1 + empcant1)) <> 0, CDbl(100), CDbl(0))
        End If
        StrSql = "UPDATE rep_comp_acum"
        StrSql = StrSql & " SET empmonto1 = empmonto1 + " & empmonto1
        StrSql = StrSql & ", empmonto2 = empmonto2 + " & empmonto2
        StrSql = StrSql & ", empcant1 = empcant1 + " & empcant1
        StrSql = StrSql & ", empcant2 = empcant2 + " & empcant2
        StrSql = StrSql & ", difmontoemp =  " & difmontoemp
        StrSql = StrSql & ", difcantemp =  " & difcantemp
        StrSql = StrSql & ", porccantemp =  " & porccantemp
        StrSql = StrSql & ", porcmontoemp =  " & porcmontoemp
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND ternro = " & Ternro
        StrSql = StrSql & " AND acunro = " & acuNro
        StrSql = StrSql & " AND concnro = " & ConcNro
        '    Flog.writeline StrSql
            objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    End If
    
    Seguir = False
    If rsConsult.EOF Then
        Seguir = True
    Else
        If rsConsult!ConcNro <> ConcNro Or rsConsult.EOF Then
            Seguir = True
        End If
    End If
    
    If Seguir Then
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
        'Actualizo las cantidad correspondientes a los conceptos en la BD
        '-------------------------------------------------------------------------------
        StrSql = "UPDATE rep_comp_acum SET concmonto1=" & concmonto1 & _
                 ",conccant1= " & conccant1 & _
                 ",concmonto2=" & concmonto2 & _
                 ",conccant2=" & conccant2 & _
                 ",difmontoconc=" & difmontoconc & _
                 ",porcmontoconc=" & porcmontoconc & _
                 ",difcantconc=" & difcantconc & _
                 ",porccantconc=" & porccantconc & _
                 " WHERE concnro = " & ConcNro & " AND acunro = " & acuNro
      '  Flog.writeline StrSql
        objConn.Execute StrSql, , adExecuteNoRecords
        
        concmonto1 = CDbl(0)
        concmonto2 = CDbl(0)
        conccant1 = CDbl(0)
        conccant2 = CDbl(0)
    End If
    
TiempoAcumulado = GetTickCount
cantRegistros = cantRegistros - 1
  
'StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalAcum - cantRegistros) * 100) / totalAcum) & _
         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
         " WHERE bpronro = " & NroProceso
     
'Flog.writeline StrSql
'objconn2.Execute StrSql, , adExecuteNoRecords
    
    
Loop

If rsConsult.State = adStateOpen Then rsConsult.Close

Exit Sub
            
MError:
    Flog.writeline " Error: " & Err & " - Descripción: " & Err.Description
    Flog.writeline " Ultimo SQL ejecutado: " & StrSql
    Flog.writeline " Ternro = " & Ternro
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub
'--------------------------------------------------------------------
' Se encarga de generar un rsAcum con los acumuladores de la comparacion
'--------------------------------------------------------------------
Sub CargarAcum(pronro1 As String, pronro2 As String, ByRef rsAcum As ADODB.Recordset)

Dim StrAcum As String

    StrAcum = "SELECT DISTINCT con_acum.acunro " & _
              " FROM cabliq " & _
              " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro " & _
              " INNER JOIN con_acum ON detliq.concnro = con_acum.concnro " & _
              " WHERE cabliq.pronro IN (" & pronro1 & "," & pronro2 & ") "
    If listaacunro <> "0" Then
        StrAcum = StrAcum & "AND con_acum.acunro IN (" & listaacunro & ") "
    End If
    StrAcum = StrAcum & "ORDER BY con_acum.acunro"
    
    'Flog.writeline StrAcum
    
    OpenRecordset StrAcum, rsAcum
End Sub

