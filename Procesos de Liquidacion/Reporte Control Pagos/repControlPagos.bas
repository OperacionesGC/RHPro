Attribute VB_Name = "repControlPagos"
Option Explicit

Global Const Version = "1.01"
Global Const FechaVersion = "21/10/2009"
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

Global repNro As Integer
Global conceptos As String
Global acumuladores As String
Global procesos As String
Global idUser As String

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
Dim historico As Boolean
Dim param
Dim ternro
Dim rsEmpl As New ADODB.Recordset
Dim I
Dim totalEmpleados
Dim cantRegistros
Dim PID As String
Dim parametros As String
Dim ArrParametros

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
    
    Nombre_Arch = PathFLog & "ReporteControlPagos" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    'Obtengo la cantidad de empledos a procesar
    StrSql = "SELECT count(*) AS total FROM batch_empleado WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    cantRegistros = CInt(objRs!total)
    totalEmpleados = cantRegistros
    
    objRs.Close
   
    Flog.writeline "Inicio Proceso de Control Pagos : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "Modificacion             : " & UltimaModificacion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       
       'Obtengo el nro de reporte
       repNro = objRs!bprcparam
       
       'Seteo los datos del reporte
       StrSql = "SELECT * FROM rep_control_pagos WHERE repnro = " & repNro
       OpenRecordset StrSql, objRs2
       
       If Not objRs2.EOF Then
          procesos = objRs2!procesos
          
          idUser = objRs2!idUser
          
          If IsNull(objRs2!conceptos) Then
             conceptos = "0"
          Else
             If Trim(objRs2!conceptos) = "" Then
                conceptos = "0"
             Else
                conceptos = objRs2!conceptos
             End If
          End If
          
          If IsNull(objRs2!acumuladores) Then
             acumuladores = "0"
          Else
             If Trim(objRs2!acumuladores) = "" Then
                acumuladores = "0"
             Else
                acumuladores = objRs2!acumuladores
             End If
          End If
           
          'Seteo los datos del reporte
          Call setearDatosReporte(objRs2!empresa)
       Else
          Flog.writeline "Error: No se encontro la cabezera del reporte"
          Exit Sub
       End If
      
       'EMPIEZA EL PROCESO

       'Obtengo los empleados sobre los que tengo que generar los recibos
       CargarEmpleados NroProceso, rsEmpl
       
       StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
       
       objConn.Execute StrSql, , adExecuteNoRecords
       
       'Genero por cada empleado un registro
       Do Until rsEmpl.EOF
          EmpErrores = False
          ternro = rsEmpl!ternro
          
          'Genero los datos del empleado
          Call generarDatosEmpleado(ternro)
                
          'Actualizo el estado del proceso
          TiempoAcumulado = GetTickCount
          
          cantRegistros = cantRegistros - 1
          
          StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
             
          objConn.Execute StrSql, , adExecuteNoRecords
          
          'Si se generaron todos los datos del empleado correctamente lo borro
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
   
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
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
' Se encarga de buscar los conceptos y acumuladores de un empleado
'--------------------------------------------------------------------
Sub generarDatosEmpleado(ByVal ternro)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset

On Error GoTo MError

'------------------------------------------------------------------
'Busco todos los conceptos del empleado
'------------------------------------------------------------------
StrSql = " SELECT concepto.concnro, concepto.concabr, detliq.dlimonto, cabliq.pronro "
StrSql = StrSql & " From cabliq"
StrSql = StrSql & " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro"
StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
StrSql = StrSql & " WHERE cabliq.pronro IN (" & procesos & ") AND"
StrSql = StrSql & "      detliq.concnro IN (" & conceptos & ") AND"
StrSql = StrSql & "      cabliq.Empleado =" & ternro

OpenRecordset StrSql, rsConsult

Do Until rsConsult.EOF
    StrSql = " INSERT INTO rep_cont_pagos_det "
    StrSql = StrSql & " (bpronro , repnro, ternro, conc_acum,"
    StrSql = StrSql & " codigo, descripcion, monto, pronro)"
    StrSql = StrSql & " VALUES "
    StrSql = StrSql & "(" & NroProceso
    StrSql = StrSql & "," & repNro
    StrSql = StrSql & "," & ternro
    StrSql = StrSql & ",0"
    StrSql = StrSql & "," & rsConsult!concnro
    StrSql = StrSql & ",'" & rsConsult!concabr & "'"
    StrSql = StrSql & "," & numberForSQL(rsConsult!dlimonto)
    StrSql = StrSql & "," & rsConsult!proNro
    StrSql = StrSql & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords

    rsConsult.MoveNext
Loop

rsConsult.Close

'------------------------------------------------------------------
'Busco todos los acumuladores del empleado
'------------------------------------------------------------------
StrSql = " SELECT acumulador.acudesabr, acu_liq.almonto, cabliq.pronro, acumulador.acunro "
StrSql = StrSql & " From cabliq"
StrSql = StrSql & " INNER JOIN acu_liq ON cabliq.cliqnro = acu_liq.cliqnro"
StrSql = StrSql & " INNER JOIN acumulador ON acumulador.acunro = acu_liq.acunro"
StrSql = StrSql & " WHERE cabliq.pronro IN (" & procesos & ") AND"
StrSql = StrSql & "       acu_liq.acunro IN (" & acumuladores & ") AND"
StrSql = StrSql & "       cabliq.Empleado =" & ternro

OpenRecordset StrSql, rsConsult

Do Until rsConsult.EOF
    StrSql = " INSERT INTO rep_cont_pagos_det "
    StrSql = StrSql & " (bpronro , repnro, ternro, conc_acum,"
    StrSql = StrSql & " codigo, descripcion, monto, pronro)"
    StrSql = StrSql & " VALUES "
    StrSql = StrSql & "(" & NroProceso
    StrSql = StrSql & "," & repNro
    StrSql = StrSql & "," & ternro
    StrSql = StrSql & ",1"
    StrSql = StrSql & "," & rsConsult!acuNro
    StrSql = StrSql & ",'" & rsConsult!acudesabr & "'"
    StrSql = StrSql & "," & numberForSQL(rsConsult!almonto)
    StrSql = StrSql & "," & rsConsult!proNro
    StrSql = StrSql & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords

    rsConsult.MoveNext
Loop

rsConsult.Close

Exit Sub

MError:
    Flog.writeline "Error en el tercero " & ternro & " Error: " & Err.Description
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


Function strForSQL(Str)
   
  If IsNull(Str) Then
     strForSQL = "NULL"
  Else
     strForSQL = Str
  End If

End Function

Sub setearDatosReporte(ByVal empresa)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim ternro
Dim profecpago
Dim EmprLogo
Dim EmprLogoAlto
Dim EmprLogoAncho
Dim emprTer

    ' -------------------------------------------------------------------------
    ' Busco los datos de la empresa
    '--------------------------------------------------------------------------
    
    StrSql = "SELECT empresa.ternro " & _
        " From estructura " & _
        " INNER JOIN empresa ON empresa.estrnro = estructura.estrnro " & _
        " WHERE empresa.estrnro = " & empresa
    
    OpenRecordset StrSql, rsConsult
    
    emprTer = 0
    
    If rsConsult.EOF Then
        Flog.writeline "No se encontró la empresa"
        Exit Sub
    Else
        emprTer = rsConsult!ternro
    End If
    
    rsConsult.Close
    
    'Consulta para buscar el logo de la empresa
    StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
        " From ter_imag " & _
        " INNER JOIN tipoimag ON tipoimag.tipimnro = 1 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
        " AND ter_imag.ternro =" & emprTer
    
    OpenRecordset StrSql, rsConsult
    If rsConsult.EOF Then
        Flog.writeline "No se encontró el Logo de la Empresa"
        EmprLogo = ""
        EmprLogoAlto = 0
        EmprLogoAncho = 0
    Else
        EmprLogo = rsConsult!tipimdire & rsConsult!terimnombre
        EmprLogoAlto = rsConsult!tipimaltodef
        EmprLogoAncho = rsConsult!tipimanchodef
    End If

    rsConsult.Close
    
    'Actualizo los datos del reporte
    StrSql = "UPDATE rep_control_pagos SET " & _
       " logo = '" & EmprLogo & "' " & _
       " ,logoalto = " & EmprLogoAlto & _
       " ,logoancho = " & EmprLogoAncho & _
       " ,bpronro = " & NroProceso & _
       " WHERE repnro = " & repNro
    
    objConn.Execute StrSql, , adExecuteNoRecords

    'Actualizo los datos de la auditoria del reporte
    StrSql = "INSERT INTO rep_cont_pagos_aud (bpronro,repnro,fecha,hora,iduser,accion) VALUES (" & _
       " " & NroProceso & _
       " ," & repNro & _
       " ," & ConvFecha(Date) & _
       " ,'" & Mid(Time(), 1, 8) & "'" & _
       " ,'" & idUser & "'" & _
       " ,'Reporte generado.'" & ")"
       
    
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub


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


