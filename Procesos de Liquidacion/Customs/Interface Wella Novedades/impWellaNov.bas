Attribute VB_Name = "impWella"
Option Explicit

Global Const Version = "1.01"
Global Const FechaModificacion = "31/07/2009"
Global Const UltimaModificacion = "MB - Encriptacion de string connection"

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

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
Global fechaDesde
Global fechaHasta
Global concTipoDia(100)
Global paramTipoDia(100)
Global concReintegros
Global paramReintegros
Global concHN
Global paramHN
Global concHB
Global paramHB
Global concHE
Global paramHE
Global concHEB
Global paramHEB
Global topeHEB
Global empresaCod

Global ObtReintegros As Boolean
Global ObtLicencias As Boolean
Global ObtHoras As Boolean
Global GenerarNov As Boolean

Global pliqNro
Global pliqdesde As Date
Global pliqhasta As Date
Global concPorcVenta
Global paramPorcVenta
Global concPorcCobr
Global paramPorcCobr


Dim topeHorasLic
Dim basicoHoras

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
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim I
Dim totalEmpleados
Dim cantRegistros
Dim PID As String
Dim ArrParametros
Dim parametros As String

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
    
    On Error GoTo CE
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    
    TiempoInicialProceso = GetTickCount
    'OpenConnection strconexion, objConn
    
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "InterfaceWella" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
     ' Obtengo el Process ID
    PID = GetCurrentProcessId
    
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaModificacion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline


    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    Flog.writeline "Inicio Interface Wella : " & Now
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
    
       'Obtengo los parametros del proceso
        Flog.writeline "Obtengo los parametros del proceso"
       parametros = objRs!bprcparam
       ArrParametros = Split(parametros, "@")
       
       pliqNro = ArrParametros(0)
       
       'Busco en el confrep las asociaciones para generar los parametros
       StrSql = " SELECT pliqdesde, pliqhasta, pliqdesc FROM periodo "
       StrSql = StrSql & " WHERE pliqnro = " & pliqNro
      
       OpenRecordset StrSql, objRs2
       
       If objRs2.EOF Then
            Flog.writeline "No existe el periodod de liquidacion " & pliqNro
            Exit Sub
       Else
        Flog.writeline "Periodo: [" & pliqNro & "] " & objRs2!pliqdesc & " (" & objRs2!pliqdesde & "-" & objRs2!pliqhasta & ")"
            pliqdesde = objRs2!pliqdesde
            pliqhasta = objRs2!pliqhasta
       End If
       objRs2.Close
       
       'EMPIEZA EL PROCESO
       
       'Busco en el confrep las asociaciones para generar los parametros
       StrSql = " SELECT * FROM confrep "
       StrSql = StrSql & " WHERE repnro = 149 "
      
       OpenRecordset StrSql, objRs2
       
       If objRs2.EOF Then
          Flog.writeline "No esta configurado el ConfRep"
          Exit Sub
       End If
       
       Flog.writeline "Obtengo los datos del confrep"
       
       concPorcVenta = 0
       paramPorcVenta = 0
       concPorcCobr = 0
       paramPorcCobr = 0
       Do Until objRs2.EOF
          
          If CLng(objRs2!confnrocol) = 1 Then
            'Concepto de Porcentaje Ventas
            If (objRs2!conftipo = "CO") Then
                Flog.writeline "Confrep columna " & objRs2!confnrocol
                concPorcVenta = buscarConcepto(objRs3, objRs2!confval2)
            End If
            'Parametro de Porcentaje Ventas
            If (objRs2!conftipo = "PAR") Then
                Flog.writeline "Confrep columna " & objRs2!confnrocol
                paramPorcVenta = objRs2!confval
            End If
          End If
          If CLng(objRs2!confnrocol) = 2 Then
            'Concepto de Porcentaje Cobranzas
            If (objRs2!conftipo = "CO") Then
                Flog.writeline "Confrep columna " & objRs2!confnrocol
                concPorcCobr = buscarConcepto(objRs3, objRs2!confval2)
            End If
            'Parametro de Porcentaje Cobranzas
            If (objRs2!conftipo = "PAR") Then
                Flog.writeline "Confrep columna " & objRs2!confnrocol
                paramPorcCobr = objRs2!confval
            End If
          End If
          
          objRs2.MoveNext
       Loop
       objRs2.Close

       
       Call generarNovedades
    
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
    
    Flog.writeline "Fin: " & Now
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

Function sinDatos(Str)

  If IsNull(Str) Then
     sinDatos = True
  Else
     If Trim(Str) = "" Then
        sinDatos = True
     Else
        sinDatos = False
     End If
  End If

End Function


Function buscarConcepto(ByRef rsConcepto As ADODB.Recordset, ByVal Conccod)

On Error GoTo MError

    'Busco el concnro del concepto
    StrSql = "SELECT concnro FROM concepto WHERE conccod = '" & Conccod & "'"
    
    Flog.writeline "Buscando Concepto: " & StrSql
    
    OpenRecordset StrSql, rsConcepto
    
    If rsConcepto.EOF Then
       buscarConcepto = 0
    Else
       buscarConcepto = rsConcepto!concnro
    End If
    
    rsConcepto.Close
    
    Exit Function
    
MError:
    Flog.writeline "Error en buscarEmpleado: " & Err.Description
    HuboErrores = True
    
End Function

Function buscarEmpleado(ByRef rsEmpleado As ADODB.Recordset, ByVal empleg)

On Error GoTo MError

    Flog.writeline "Buscando si se encuentra al empleado " & empleg

    'Busco el ternro del empleado
    StrSql = "SELECT ternro FROM empleado WHERE empleg = " & empleg
    
    Flog.writeline "Buscando Empleado: " & StrSql
    
    OpenRecordset StrSql, rsEmpleado
    
    If rsEmpleado.EOF Then
       buscarEmpleado = 0
    Else
       buscarEmpleado = rsEmpleado!ternro
    End If
    
    rsEmpleado.Close
    
    Exit Function
    
MError:
    Flog.writeline "Error en buscarEmpleado: " & Err.Description
    HuboErrores = True

End Function

Function buscarTipoLicencia(ByRef rsTipoLic As ADODB.Recordset, ByVal codigo)

On Error GoTo MError

    Flog.writeline "Buscando el tipo de licencia:" & codigo

    'Busco el tdnro de la licencia
    StrSql = "SELECT tdnro FROM tipdia WHERE tdsigla LIKE '" & Mid(codigo, 1, 2) & "%'"
    
    Flog.writeline "Buscando Tipo Licencia: " & StrSql
    
    OpenRecordset StrSql, rsTipoLic
    
    If rsTipoLic.EOF Then
       buscarTipoLicencia = 0
    Else
       buscarTipoLicencia = rsTipoLic!tdnro
    End If
    
    rsTipoLic.Close
    
    Exit Function
    
MError:
    Flog.writeline "Error en buscarTipoLicencia: " & Err.Description
    HuboErrores = True

End Function
                 




Function numberForSQL(Str)
   
  numberForSQL = Replace(Str, ",", ".")

End Function

Function fechaToXML(Fecha)
  
  Fecha = CDate(Fecha)
  
  fechaToXML = Month(Fecha) & "/" & Day(Fecha) & "/" & Year(Fecha)

End Function

Sub generarNovedades()

Dim rsEmpleados As New ADODB.Recordset
Dim rsFactura As New ADODB.Recordset
Dim rsTrab As New ADODB.Recordset
Dim tieneLic As Boolean
Dim tieneLicVac As Boolean
Dim tieneLicOtras As Boolean
Dim horasLic
Dim horasTrab
Dim Valor
Dim valorAux
Dim totalReg
Dim I

Dim TotalPorcVenta
Dim TotalPorcCobr
Dim Debito

On Error GoTo MError

   Flog.writeline "Empezando a generar las novedades para los empleados"
   
    'Actualizo el progreso
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = 75 " & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
             " WHERE bpronro = " & NroProceso
             
    objConn.Execute StrSql, , adExecuteNoRecords

   StrSql = " SELECT DISTINCT empleado.ternro, empleg, terape, ternom FROM empleado " & _
            " INNER JOIN wella_zonas_emp ON empleado.ternro = wella_zonas_emp.ternro "
            
   OpenRecordset StrSql, rsEmpleados

   totalReg = rsEmpleados.RecordCount
   I = 0
   Do Until rsEmpleados.EOF
      'Busco para cada empleado las facturas asociadas
      Flog.writeline "=================="
      Flog.writeline "Generando las novedades para el empleado " & rsEmpleados!empleg & " - " & rsEmpleados!terape & ", " & rsEmpleados!ternom

      StrSql = "SELECT wella_zonas_emp.zonnro, factcomprnro, porc_vta, porc_cob, factimporte, codmovdebito, codmovcod " & _
               " FROM wella_zonas_emp " & _
               " INNER JOIN wella_facturacion ON wella_zonas_emp.zonnro = wella_facturacion.zonnro " & _
               " INNER JOIN wella_cod_mov ON wella_facturacion.codmovnro = wella_cod_mov.codmovnro " & _
               " WHERE wella_zonas_emp.ternro = " & rsEmpleados!ternro & _
               " AND wella_cod_mov.codmovafecta = -1 " & _
               " ORDER BY wella_zonas_emp.zonnro "
      
      OpenRecordset StrSql, rsFactura
      
      TotalPorcVenta = 0
      TotalPorcCobr = 0
      Do Until rsFactura.EOF
            Flog.writeline "# Compr Nro: " & rsFactura!factcomprnro & _
                            " # " & IIf((rsFactura!codmovdebito = -1), "Debito", "Credito") & _
                            " # Mov: " & rsFactura!codmovcod & _
                            " # Imp: " & rsFactura!factimporte & _
                            " # Vta: " & rsFactura!porc_vta & "%" & _
                            " # Cob: " & rsFactura!porc_cob & "%"

            Debito = IIf((rsFactura!codmovdebito = -1), -1, 1)
            TotalPorcVenta = TotalPorcVenta + Debito * (rsFactura!factimporte * rsFactura!porc_vta / 100)
            TotalPorcCobr = TotalPorcCobr + Debito * (rsFactura!factimporte * rsFactura!porc_cob / 100)
            rsFactura.MoveNext
            If rsFactura.EOF Then
                Call generarNovedad(rsEmpleados!ternro, concPorcVenta, paramPorcVenta, TotalPorcVenta)
                Call generarNovedad(rsEmpleados!ternro, concPorcCobr, paramPorcCobr, TotalPorcCobr)
            End If
      Loop
   
      'Actualizo el estado del proceso
      TiempoAcumulado = GetTickCount
      StrSql = "UPDATE batch_proceso SET bprcprogreso = " & numberForSQL(CStr((((I + 1) * 25) / totalReg) + 75)) & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                 " WHERE bpronro = " & NroProceso
                 
      objConn.Execute StrSql, , adExecuteNoRecords
   
      rsEmpleados.MoveNext
      I = I + 1
   Loop
   
   rsEmpleados.Close
   
   Set rsEmpleados = Nothing
   Set rsFactura = Nothing
'   Set rsTrab = Nothing
   
   Flog.writeline "Termina de generar las novedades"
   
   Exit Sub

MError:
    Flog.writeline "Error en generarNovedades, Error: " & Err.Description
    Exit Sub

End Sub




Sub generarNovedad(ternro, concnro, tpanro, Valor)

Dim rsConsulta As New ADODB.Recordset

On Error GoTo MError

    If concnro <> 0 And tpanro <> 0 Then
    
        StrSql = " SELECT * FROM novemp WHERE empleado = " & ternro
        StrSql = StrSql & " AND concnro = " & concnro
        StrSql = StrSql & " AND tpanro = " & tpanro
        StrSql = StrSql & " AND ( nedesde = " & ConvFecha(pliqdesde)
        StrSql = StrSql & " OR nehasta = " & ConvFecha(pliqhasta)
        StrSql = StrSql & " ) "
        
        OpenRecordset StrSql, rsConsulta
        
        If rsConsulta.EOF Then
        
           StrSql = "INSERT INTO novemp "
           StrSql = StrSql & "(empleado, concnro, tpanro, nevalor, nevigencia, nedesde, nehasta, neretro, nepliqdesde, nepliqhasta)"
           StrSql = StrSql & " values (" & ternro & ", " & concnro & ", " & tpanro & ", " & numberForSQL(Valor) & ", "
           StrSql = StrSql & "-1," & ConvFecha(pliqdesde) & "," & ConvFecha(pliqhasta) & ",null,null,null) "
          
        Else
        
           StrSql = "UPDATE novemp "
'           StrSql = StrSql & " SET nevalor = nevalor + " & numberForSQL(Valor)
           StrSql = StrSql & " SET nevalor = " & numberForSQL(Valor)
           StrSql = StrSql & " WHERE nenro = " & rsConsulta!nenro
        
        End If
        
        rsConsulta.Close
        
        Flog.writeline "Generando novedad: " & StrSql
    
        objConn.Execute StrSql, , adExecuteNoRecords
    
    End If
    
    Set rsConsulta = Nothing
    
Exit Sub

MError:
    Flog.writeline "Error en generarNovedades para el ternro: " & ternro & " Error: " & Err.Description
    Exit Sub

End Sub

