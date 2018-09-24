Attribute VB_Name = "repNomMensxEmpl"
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

Global pliqnro1 As Integer
Global pliqdesc1 As String
Global pliqmesanio1 As String
Global pronro1 As String
Global pliqnro2 As Integer
Global pliqdesc2 As String
Global pliqmesanio2 As String
Global pronro2 As String
Global listaacunro As String
Global vol_cod As Integer
Global masinro As Integer
Global listaConcAcumNeto As String
Global Tabulador As Long


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
Dim rsEmpl As New ADODB.Recordset
Dim i
Dim totalAcum
Dim cantRegistros
Dim PID As String
Dim parametros As String
Dim ArrParametros

Dim ternro As Integer
Dim empleg As String
Dim terape As String
Dim ternom2 As String
Dim terape2 As String
Dim ternom As String

Dim vol_desc As String
Dim masidesc As String

'Dim ArrParametros

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
    
    Nombre_Arch = PathFLog & "ReporteCompAcum" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "Inicio Proceso de Comparación de Acumuladores : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
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
       
       'Obtengo el proceso de volcado
       vol_cod = ArrParametros(5)
       
       'Obtengo el modelo de asiento
       masinro = ArrParametros(6)
       
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
       
       'Busco el modelo de Asiento
       StrSql = "SELECT masidesc FROM mod_asiento WHERE masinro = " & masinro
       OpenRecordset StrSql, objRs
        
       masidesc = ""
       If Not objRs.EOF Then
          masidesc = objRs!masidesc
       End If
        
       If objRs.State = adStateOpen Then objRs.Close
       
       'Busco el Proceso de Volcado
       StrSql = "SELECT vol_desc FROM proc_vol WHERE vol_cod = " & vol_cod
       OpenRecordset StrSql, objRs
        
       vol_desc = ""
       If Not objRs.EOF Then
          vol_desc = objRs!vol_desc
       End If
        
       If objRs.State = adStateOpen Then objRs.Close
       
       StrSql = "INSERT INTO rep_men_empl (bpronro,pliqnro1,pliqdesc1,pliqmesanio1,pliqnro2,pliqdesc2," & _
                "pliqmesanio2,vol_cod,vol_desc,masinro,masidesc) " & _
                "VALUES (" & _
                 NroProceso & "," & pliqnro1 & ",'" & pliqdesc1 & "','" & pliqmesanio1 & "'," & _
                 pliqnro2 & ",'" & pliqdesc2 & "','" & pliqmesanio2 & "'," & vol_cod & ",'" & _
                 vol_desc & "'," & masinro & ",'" & masidesc & "')"
            
       objConn.Execute StrSql, , adExecuteNoRecords
       
       
       ' Este es un caso particular para el unico acumulador que posee conceptos (acunro = 6 - NETO)
       ' Busco los conceptos asociados al acumulador
       StrSql = "SELECT concnro FROM con_acum WHERE acunro = 6"
       OpenRecordset StrSql, objRs
    
       listaConcAcumNeto = ","
       Do Until objRs.EOF
          listaConcAcumNeto = listaConcAcumNeto & objRs!concnro & ","
          objRs.MoveNext
       Loop
        
       If objRs.State = adStateOpen Then objRs.Close
       
       
       'Obtengo los acumuladores sobre los que deberia generar la comparación
       Call CargarEmpleados(pronro1, pronro2, rsEmpl)
       
       cantRegistros = CInt(rsEmpl.RecordCount)
       totalAcum = CInt(cantRegistros)
       
       ' Genero los datos
       Do Until rsEmpl.EOF
            EmpErrores = False
            ternro = rsEmpl!ternro
              
            ' Genero el procesamiento del acumulador
            Flog.writeline "Generando datos Empleado " & rsEmpl!empleg
                      
            ' Genero los datos del Empleado
            ternro = rsEmpl!ternro
            empleg = rsEmpl!empleg
            terape = ""
            If Not IsNull(rsEmpl!terape) Then
                terape = rsEmpl!terape
            End If
            ternom2 = ""
            If Not IsNull(rsEmpl!ternom2) Then
                ternom2 = rsEmpl!ternom2
            End If
            terape2 = ""
            If Not IsNull(rsEmpl!terape2) Then
                terape2 = rsEmpl!terape2
            End If
            ternom = ""
            If Not IsNull(rsEmpl!ternom) Then
                ternom = rsEmpl!ternom
            End If
            Call GenerarDatosEmpleado(ternro, pronro1, pronro2, masinro, vol_cod, empleg, terape, ternom2, terape2, ternom)
            
            TiempoAcumulado = GetTickCount
              
            cantRegistros = cantRegistros - 1
              
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & ((totalAcum - cantRegistros) * 100) / totalAcum & _
                     ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                     " WHERE bpronro = " & NroProceso
                 
            objConn.Execute StrSql, , adExecuteNoRecords
             
            rsEmpl.MoveNext
       Loop
       
       If rsEmpl.State = adStateOpen Then rsEmpl.Close
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
' Se encarga de generar la comparacion para el acumulador
'--------------------------------------------------------------------
Sub GenerarDatosEmpleado(ByVal ternro As Long, ByVal pronro1 As String, ByVal pronro2 As String, ByVal masinro As Integer, ByVal vol_cod As Integer, ByVal empleg As String, ByVal terape As String, ByVal ternom2 As String, ByVal terape2 As String, ByVal ternom As String)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim rsConsult2 As New ADODB.Recordset

'Variables donde se guardan los datos del INSERT final
Dim empmonto1 As Double
Dim empcant1 As Double
Dim empmonto2 As Double
Dim empcant2 As Double
Dim difmontoemp As Double
Dim porcmontoemp As Double
Dim difcantemp As Double
Dim porccantemp As Double

Dim concnro As Integer
Dim tconcepto As Integer
Dim banco As String
Dim pers_sub_area As String
Dim pers_area As String
Dim centrocosto As String
Dim sucursal As String
Dim ctabancaria As String
Dim ctadebito As String
Dim ctacredito As String
Dim Conccod As String
Dim concabr As String


On Error GoTo MError

'------------------------------------------------------------------
'Busco el valor del centro de costo
'------------------------------------------------------------------
StrSql = " SELECT estrcodext "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htethasta IS NULL AND his_estructura.tenro=5 AND his_estructura.ternro=" & ternro
       
OpenRecordset StrSql, rsConsult
centrocosto = ""
If Not rsConsult.EOF Then
   centrocosto = rsConsult!estrcodext
End If
If rsConsult.State = adStateOpen Then rsConsult.Close

'------------------------------------------------------------------
'Busco el valor del Personal Area
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htethasta IS NULL AND his_estructura.tenro=49 AND his_estructura.ternro=" & ternro
       
OpenRecordset StrSql, rsConsult
pers_area = ""
If Not rsConsult.EOF Then
   pers_area = rsConsult!estrdabr
End If
If rsConsult.State = adStateOpen Then rsConsult.Close

'------------------------------------------------------------------
'Busco el valor del Personal Sub-Area
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htethasta IS NULL AND his_estructura.tenro=48 AND his_estructura.ternro=" & ternro
       
OpenRecordset StrSql, rsConsult
pers_sub_area = ""
If Not rsConsult.EOF Then
   pers_sub_area = rsConsult!estrdabr
End If
If rsConsult.State = adStateOpen Then rsConsult.Close

'------------------------------------------------------------------
'Busco el valor de la Sucursal
'------------------------------------------------------------------
StrSql = " SELECT estrcodext "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htethasta IS NULL AND his_estructura.tenro=1 AND his_estructura.ternro=" & ternro
       
OpenRecordset StrSql, rsConsult
sucursal = ""
If Not rsConsult.EOF Then
   sucursal = rsConsult!estrcodext
End If
If rsConsult.State = adStateOpen Then rsConsult.Close

'------------------------------------------------------------------
'Busco el valor de Número Cuenta
'------------------------------------------------------------------
StrSql = " SELECT ctabnro, ctabcbu, estructura.estrdabr "
StrSql = StrSql & " From ctabancaria"
StrSql = StrSql & " LEFT JOIN banco ON ctabancaria.banco = banco.ternro "
StrSql = StrSql & " LEFT JOIN estructura ON estructura.estrnro = banco.estrnro "
StrSql = StrSql & " WHERE ctabestado = -1 AND ctabancaria.ternro=" & ternro
       
OpenRecordset StrSql, rsConsult
ctabancaria = ""
banco = ""
If Not rsConsult.EOF Then
    banco = rsConsult!estrdabr
    If rsConsult!ctabcbu <> "" Then
        ctabancaria = rsConsult!ctabcbu
    Else
        ctabancaria = rsConsult!ctabnro
    End If
End If
If rsConsult.State = adStateOpen Then rsConsult.Close

'------------------------------------------------------------------
' Busco los conceptos en la liquidacion
'------------------------------------------------------------------
StrSql = "SELECT detliq.concnro, cabliq.empleado, proceso.pliqnro, SUM(detliq.dlimonto) AS sumdlimonto, " & _
         "SUM(detliq.dlicant) AS sumdlicant, conccod, concabr, concepto.tconnro " & _
         "FROM proceso " & _
         "INNER JOIN cabliq ON proceso.pronro = cabliq.pronro " & _
         "INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro " & _
         "LEFT JOIN concepto ON detliq.concnro = concepto.concnro " & _
         "WHERE cabliq.pronro IN (" & pronro1 & "," & pronro2 & ") " & _
         "AND cabliq.empleado = " & ternro & " " & _
         "GROUP BY proceso.pliqnro, cabliq.empleado, detliq.concnro, conccod, concabr, concepto.tconnro " & _
         "ORDER BY detliq.concnro, cabliq.empleado"

OpenRecordset StrSql, rsConsult

empmonto1 = 0
empcant1 = 0
empmonto2 = 0
empcant2 = 0
If Not rsConsult.EOF Then
    concnro = rsConsult!concnro
End If
Do Until rsConsult.EOF
    
    If rsConsult!concnro <> concnro Then
        
        tconcepto = CalcularMapeo(Conccod, "CONCEPTOS", 0)
        
        '------------------------------------------------------------------
        'Busco el valor de Cuenta contable (Débito o Crédito)
        '------------------------------------------------------------------
        StrSql = " SELECT linea_asi.cuenta, linea_asi.dh "
        StrSql = StrSql & " FROM detalle_asi"
        StrSql = StrSql & " INNER JOIN linea_asi ON detalle_asi.masinro = linea_asi.masinro AND"
        StrSql = StrSql & " detalle_asi.vol_cod = linea_asi.vol_cod AND detalle_asi.cuenta = linea_asi.cuenta"
        StrSql = StrSql & " WHERE tipoorigen = 1 AND detalle_asi.masinro = " & masinro & " AND ternro=" & ternro
        StrSql = StrSql & " AND detalle_asi.vol_cod = " & vol_cod & " AND origen = " & concnro
        
        OpenRecordset StrSql, rsConsult2
        ctadebito = ""
        ctacredito = ""
        Do Until rsConsult2.EOF
            If CInt(rsConsult2!dh) = -1 Then
                ctadebito = rsConsult2!cuenta
            End If
                    
            If CInt(rsConsult2!dh) = 0 Then
                ctacredito = rsConsult2!cuenta
            End If
            rsConsult2.MoveNext
        Loop
        If rsConsult2.State = adStateOpen Then rsConsult2.Close
        
        ' Este es un caso particular para el unico acumulador que posee conceptos (acunro = 6 - NETO)
        ' Para este caso la cuenta de credito se cablea
        If InStr(listaConcAcumNeto, "," & concnro & ",") > 0 Then
            ctadebito = "21220101"
        End If
        
        If (empmonto1 <> 0 Or empcant1 <> 0 Or empmonto2 <> 0 Or empcant2 <> 0) Then
            '------------------------------------------------------------------
            ' Realizo los calculos sobre los campos
            '------------------------------------------------------------------
            difmontoemp = CDbl(empmonto1) - CDbl(empmonto2)
            If CDbl(empmonto2) <> 0 Then
                porcmontoemp = (CDbl(difmontoemp) * CDbl(100)) / CDbl(empmonto2)
            Else
                porcmontoemp = IIf(CDbl(empmonto1) <> 0, CDbl(100), CDbl(0))
            End If
            difcantemp = CDbl(empcant1) - CDbl(empcant2)
            If CDbl(empcant2) <> 0 Then
                porccantemp = (CDbl(difcantemp) * CDbl(100)) / CDbl(empcant2)
            Else
                porccantemp = IIf(CDbl(empcant1) <> 0, CDbl(100), CDbl(0))
            End If
                
            '-------------------------------------------------------------------------------
            'Inserto los datos en la BD
            '-------------------------------------------------------------------------------
            StrSql = "INSERT INTO rep_men_empl_det(bpronro,ternro,empleg,terape,ternom2,terape2," & _
                     "ternom,centrocosto,pers_area,pers_sub_area,banco,sucursal,ctabancaria," & _
                     "concnro,conccod,concabr,empmonto1,empcant1,empmonto2,empcant2,difmontoemp," & _
                     "porcmontoemp,difcantemp,porccantemp,tconcepto,ctadebito,ctacredito) "
            StrSql = StrSql & "VALUES (" & _
                     NroProceso & "," & ternro & ",'" & empleg & "','" & _
                     terape & "','" & ternom2 & "','" & terape2 & "','" & ternom & "','" & _
                     centrocosto & "','" & Trim(pers_area) & "','" & Trim(pers_sub_area) & "','" & Trim(banco) & "','" & _
                     sucursal & "','" & Trim(ctabancaria) & "'," & concnro & ",'" & Trim(Conccod) & "','" & _
                     concabr & "'," & empmonto1 & "," & empcant1 & "," & empmonto2 & "," & empcant2 & "," & _
                     difmontoemp & "," & porcmontoemp & "," & difcantemp & "," & porccantemp & "," & _
                     tconcepto & ",'" & Trim(ctadebito) & "','" & Trim(ctacredito) & "')"
                
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        empmonto1 = 0
        empcant1 = 0
        empmonto2 = 0
        empcant2 = 0
        
    End If
    
    If rsConsult!pliqnro = pliqnro1 Then
        If Not IsNull(rsConsult!sumdlimonto) Then
            empmonto1 = CDbl(rsConsult!sumdlimonto)
        End If
        If Not IsNull(rsConsult!sumdlicant) Then
            empcant1 = CDbl(rsConsult!sumdlicant)
        End If
    End If

    If rsConsult!pliqnro = pliqnro2 Then
        If Not IsNull(rsConsult!sumdlimonto) Then
            empmonto2 = CDbl(rsConsult!sumdlimonto)
        End If
        If Not IsNull(rsConsult!sumdlicant) Then
            empcant2 = CDbl(rsConsult!sumdlicant)
        End If
    End If
    
    concnro = rsConsult!concnro
    Conccod = rsConsult!Conccod
    concabr = rsConsult!concabr
    
    rsConsult.MoveNext
    
Loop

If rsConsult.State = adStateOpen Then rsConsult.Close

Exit Sub
            
MError:
    Flog.writeline " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub

'--------------------------------------------------------------------
' Se encarga de buscar las cabezas de liquidacion de los empleados que respetan el filtro
'--------------------------------------------------------------------
Sub CargarEmpleados(pronro1 As String, pronro2 As String, ByRef rsEmpl As ADODB.Recordset)

Dim StrSql As String

    StrSql = "SELECT DISTINCT empleado.ternro, empleg, terape, terape2, ternom, ternom2 " & _
             "FROM cabliq " & _
             "INNER JOIN empleado ON cabliq.empleado = empleado.ternro " & _
             "INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro " & _
             "INNER JOIN con_acum ON detliq.concnro = con_acum.concnro " & _
             "WHERE cabliq.pronro IN (" & pronro1 & "," & pronro2 & ") "
    If listaacunro <> "0" Then
        StrSql = StrSql & "AND con_acum.acunro IN (" & listaacunro & ") "
    End If
    StrSql = StrSql & "ORDER BY empleado.ternro"
    
    OpenRecordset StrSql, rsEmpl
End Sub

'----------------------------------------------------------------
'Busca cual es el mapeo de un codigo RHPro a un codigo SAP
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
            Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo para la tabla " & Tabla & " con el codigo interno " & Parametro
        End If
        
        rs_Consult.Close
    Else
        Flog.writeline Espacios(Tabulador * 3) & "Parametro incorrecto al calcular el mapero de la tabla " & Tabla
    End If
    
    CalcularMapeo = Salida

End Function

