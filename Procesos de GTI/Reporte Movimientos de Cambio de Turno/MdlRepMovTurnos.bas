Attribute VB_Name = "MdlRepMovTurnos"
Option Explicit

'Const Version = "1.0"   'primera version por Nicolas Martinez
'Const FechaVersion = "19/01/2011"

Const Version = "2.0"   'Cambios en la especificacion.
Const FechaVersion = "12/04/2011"
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Public Type TDia
    Fecha As Date
    HoraIni As String
    horafin As String
    descanso As Integer
    eslaborable As Boolean
End Type

Global semana1(1 To 7) As TDia
Global semana2(1 To 7) As TDia
Global cambios(1 To 7) As Boolean

Dim fs, f

Dim Progreso As Single

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

'Datos Globales y comunes para todos los empleados

Dim fechadesde
Dim fechahasta
Dim conmov

Dim tipodepago

Dim empresa ' se actualiza para cada emp
Dim emplegajo ' se actualiza para cada emp
Dim usuario
Dim fechabatch

Global SinError As Boolean
Private Sub Main()

Dim NombreArchivo As String
Dim Directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim StrSql2 As String
Dim SqlEmpleado As String

Dim objRs As New ADODB.Recordset
Dim objRs2 As New ADODB.Recordset
Dim objAcu As New ADODB.Recordset
Dim objRsEmpleado As New ADODB.Recordset

Dim Ternro

Dim arr

'variables mias
Dim todos

Dim empestrnro As Long
Dim empresarazon

Dim terape As String
Dim terape2 As String
Dim ternom As String
Dim ternom2 As String
Dim nomemp As String

Dim tipotarjeta As String
Dim nrotarjeta As String

Dim numaccion
Dim repid
Dim fefectiva

Dim ccosto As String
Dim puesto As String
Dim tpago As String

Dim TotalEmpleados As Long
Dim IncPorc
Dim TiempoInicialProceso
Dim TiempoAcumulado
Dim PID As String
Dim bprcparam As String

Dim puntero As Date
Dim dias As Integer
Dim Dia As TDia

Dim i As Integer
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
    
    'Obtiene los datos de como esta configurado el servidor actualmente
    Call ObtenerConfiguracionRegional
    
    Nombre_Arch = PathFLog & "Rep_MovTurnosGTI" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    OpenConnection strconexion, objConnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error GoTo CE
    
    TiempoInicialProceso = GetTickCount
    
    HuboErrores = False
        
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Numero, separador decimal    : " & NumeroSeparadorDecimal
    Flog.writeline "Numero, separador de miles   : " & NumeroSeparadorMiles
    Flog.writeline "Moneda, separador decimal    : " & MonedaSeparadorDecimal
    Flog.writeline "Moneda, separador de miles   : " & MonedaSeparadorMiles
    Flog.writeline "Formato de Fecha del Servidor: " & FormatoDeFechaCorto
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 289 AND bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
 
    If Not objRs.EOF Then
       'Obtengo los parametros del proceso
        
        'recibo el usuario pero voy y busco el nombre completo del usuario
        usuario = objRs!iduser
        StrSql = " SELECT usrnombre FROM user_per "
        StrSql = StrSql & " WHERE iduser = '" & usuario & "'"
        If objRs2.State = adStateOpen Then objRs2.Close
        OpenRecordset StrSql, objRs2
        If Not objRs2.EOF Then
            usuario = objRs2!usrnombre
        End If

        fechabatch = objRs!bprcfecha
        
        Flog.writeline
        arr = Split(objRs!bprcparam, "@")
      
        fechadesde = arr(0)
        Flog.writeline "Parametro Fecha Desde = " & fechadesde
        fechahasta = arr(1)
        Flog.writeline "Parametro Fecha Hasta = " & fechahasta
        todos = arr(2) 'Si true, se consideran todos los empleados activos que no sean de agencia, de lo contrario se mira en batch_empleado
        If todos = -1 Then
            Flog.writeline "Parametro Empleados = Todos"
        Else
            Flog.writeline "Parametro Empleados = 1 o filtro"
        End If
        conmov = arr(3)
        If conmov = -1 Then
            Flog.writeline "Unicamente se generarán datos de empleados que presenten cambios en sus turnos"
        Else
            Flog.writeline "Se generarán datos de empleados que presenten o no cambios en sus turnos"
        End If
              
        tipodepago = -1
        
        StrSql = " SELECT confnrocol,conftipo,confval FROM confrep "
        StrSql = StrSql & " WHERE repnro = 304 and confval is not null "
        StrSql = StrSql & " ORDER BY confnrocol ASC"
        If objRs2.State = adStateOpen Then objRs2.Close
        OpenRecordset StrSql, objRs2
        Flog.writeline
        Flog.writeline "Obtengo la configuracion del reporte de CONFREP"
        Flog.writeline
        If Not objRs2.EOF Then
            Do Until objRs2.EOF
             If (objRs2!confnrocol = 1) Then
                    If UCase(objRs2!conftipo) = "TP" Then
                        tipodepago = objRs2!confval
                    Else
                        Flog.writeline "Se esperaba una columna de tipo TP en la columna 1"
                    End If
            End If
            objRs2.MoveNext
            Loop
        Else
           Flog.writeline "No estan configuradas las columnas del reporte (304)"
           Exit Sub
        End If
        If objRs2.State = adStateOpen Then objRs2.Close
                          
        If (Not okConfRep()) Then
            Flog.writeline
            Flog.writeline "Fin : " & Now
            Flog.Close
            objConn.Close
            Exit Sub
        End If
                                                               
        Flog.writeline
        Flog.writeline "Configuracion de CONFREP cargada correctamente"
        Flog.writeline
        
        HuboErrores = False
                  
        'COMIENZA EL PROCESO
        
        Flog.writeline "COMIENZA EL PROCESO"
        Flog.writeline
        
        'Obtengo la lista de empleados sobre los que voy a trabajar
        
        If todos = -1 Then
            StrSql = "select empleado.empleg legajo, empleado.ternro ternro, terape, terape2, ternom,"
            StrSql = StrSql & " ternom2, empresa.ternro empternro, empresa.estrnro empestrnro, empnom,"
            StrSql = StrSql & " puedesc, estrdabr cencosto " ', tptrdes,hstjnrotar "
            StrSql = StrSql & " from empleado inner join his_estructura on his_estructura.ternro = empleado.ternro "
            StrSql = StrSql & " and his_estructura.tenro = 10 and his_estructura.htetdesde <= " & ConvFecha(fechadesde) & " and "
            StrSql = StrSql & " (his_estructura.htethasta is NULL or his_estructura.htethasta >= " & ConvFecha(fechadesde) & ")"
            StrSql = StrSql & " inner join empresa on empresa.estrnro = his_estructura.estrnro "
            StrSql = StrSql & " inner join his_estructura pues on pues.ternro = empleado.ternro "
            StrSql = StrSql & " and pues.tenro = 4 and pues.htetdesde <= " & ConvFecha(fechadesde) & " and "
            StrSql = StrSql & " (pues.htethasta is NULL or pues.htethasta >= " & ConvFecha(fechadesde) & ")"
            StrSql = StrSql & " left join puesto on puesto.estrnro = pues.estrnro "
            
            StrSql = StrSql & " inner join his_estructura ccost on ccost.ternro = empleado.ternro "
            StrSql = StrSql & " and ccost.tenro = 5 and ccost.htetdesde <= " & ConvFecha(fechadesde) & " and "
            StrSql = StrSql & " (ccost.htethasta is NULL or ccost.htethasta >= " & ConvFecha(fechadesde) & ")"
            StrSql = StrSql & " left join estructura on estructura.estrnro = ccost.estrnro "
            
'            StrSql = StrSql & " inner join gti_histarjeta on gti_histarjeta.ternro = empleado.ternro "
'            StrSql = StrSql & " inner join gti_tiptar on gti_tiptar.tptrnro = gti_histarjeta.tptrnro "
            
            StrSql = StrSql & " where empleado.empest = -1 "
            StrSql = StrSql & " ORDER BY empresa.estrnro ASC"
        Else
            StrSql = "select empleado.empleg legajo, empleado.ternro ternro, terape, terape2, ternom,"
            StrSql = StrSql & " ternom2, empresa.ternro empternro, empresa.estrnro empestrnro, empnom,"
            StrSql = StrSql & " puedesc, estrdabr cencosto " ', tptrdes,hstjnrotar "
            StrSql = StrSql & " from empleado inner join batch_empleado on batch_empleado.ternro = empleado.ternro and batch_empleado.bpronro = " & NroProceso
            StrSql = StrSql & " inner join his_estructura on his_estructura.ternro = empleado.ternro "
            StrSql = StrSql & " and his_estructura.tenro = 10 and his_estructura.htetdesde <= " & ConvFecha(fechadesde) & " and "
            StrSql = StrSql & " (his_estructura.htethasta is NULL or his_estructura.htethasta >= " & ConvFecha(fechadesde) & ")"
            StrSql = StrSql & " inner join empresa on empresa.estrnro = his_estructura.estrnro "
            StrSql = StrSql & " inner join his_estructura pues on pues.ternro = empleado.ternro "
            StrSql = StrSql & " and pues.tenro = 4 and pues.htetdesde <= " & ConvFecha(fechadesde) & " and "
            StrSql = StrSql & " (pues.htethasta is NULL or pues.htethasta >= " & ConvFecha(fechadesde) & ")"
            StrSql = StrSql & " left join puesto on puesto.estrnro = pues.estrnro "
            
            StrSql = StrSql & " inner join his_estructura ccost on ccost.ternro = empleado.ternro "
            StrSql = StrSql & " and ccost.tenro = 5 and ccost.htetdesde <= " & ConvFecha(fechadesde) & " and "
            StrSql = StrSql & " (ccost.htethasta is NULL or ccost.htethasta >= " & ConvFecha(fechadesde) & ")"
            StrSql = StrSql & " left join estructura on estructura.estrnro = ccost.estrnro "
            
    '        StrSql = StrSql & " inner join gti_histarjeta on gti_histarjeta.ternro = empleado.ternro "
    '        StrSql = StrSql & " inner join gti_tiptar on gti_tiptar.tptrnro = gti_histarjeta.tptrnro "

            StrSql = StrSql & " ORDER BY empresa.estrnro ASC"
        End If
        
        OpenRecordset StrSql, objRsEmpleado
        
        TotalEmpleados = objRsEmpleado.RecordCount
        
        Flog.writeline "Cantidad de Empleados a Procesar: " & TotalEmpleados
        Flog.writeline
        
        If TotalEmpleados <> 0 Then
            IncPorc = (100 / TotalEmpleados)
        Else
            IncPorc = 100
        End If
        Progreso = 0
        
        empestrnro = 0
              
        TiempoAcumulado = GetTickCount
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(TotalEmpleados) & "' WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
               
        Do Until objRsEmpleado.EOF
        
            SinError = True
                    
            If objRsEmpleado!empestrnro <> empestrnro Then
                empresarazon = objRsEmpleado!empnom
                empestrnro = objRsEmpleado!empestrnro
                empresa = empestrnro
                Flog.writeline "Empleados de la empresa " & empresa & " ( " & objRsEmpleado!empnom & " )"
                Flog.writeline
            End If
            
            Flog.writeline Espacios(Tabulador) & "Inicio Empleado: " & objRsEmpleado!Legajo
            
            Ternro = objRsEmpleado!Ternro
            emplegajo = objRsEmpleado!Legajo
            
            'Nombres y apellidos del empleado
            Flog.writeline Espacios(Tabulador * 2) & "Recuperando datos personales del empleado"
            
            terape = IIf(IsNull(objRsEmpleado!terape), "", objRsEmpleado!terape)
            terape2 = IIf(IsNull(objRsEmpleado!terape2), "", objRsEmpleado!terape2)
            ternom = IIf(IsNull(objRsEmpleado!ternom), "", objRsEmpleado!ternom)
            ternom2 = IIf(IsNull(objRsEmpleado!ternom2), "", objRsEmpleado!ternom2)
            
            nomemp = ternom
            If ternom2 <> "" Then
                nomemp = nomemp & " " & ternom2
            End If
            nomemp = nomemp & " " & terape
            If terape2 <> "" Then
                nomemp = nomemp & " " & terape2
            End If
                                   
            Flog.writeline Espacios(Tabulador * 2) & "Recuperando datos referidos al puesto y centro de costo"
            
            ccosto = objRsEmpleado!cencosto
            puesto = objRsEmpleado!puedesc
            
            
            StrSql = " select estrdabr from his_estructura inner join estructura on his_estructura.estrnro = estructura.estrnro "
            StrSql = StrSql & " and his_estructura.tenro = " & tipodepago & " and his_estructura.ternro = " & Ternro & " AND "
            StrSql = StrSql & " his_estructura.htetdesde <= " & ConvFecha(fechadesde) & " and "
            StrSql = StrSql & " (his_estructura.htethasta is NULL or his_estructura.htethasta >= " & ConvFecha(fechadesde) & ")"
            If objRs2.State = adStateOpen Then objRs2.Close
            OpenRecordset StrSql, objRs2
            If Not objRs2.EOF Then
                tpago = objRs2!estrdabr
            Else
                Flog.writeline Espacios(Tabulador * 2) & "No se encontro informacion de tipo de pago para este empleado"
                tpago = ""
            End If
                      
            'me fijo si el empleado tiene cambios en los turnos
            
'            StrSql = " SELECT gti_notifhor_det.notifnro numaccion, wd, wdlab, wddesde, wdhasta, wdpausa,gti_notifhor.fechadesde fdesde "
'            StrSql = StrSql & " FROM gti_notifhor INNER JOIN gti_notifhor_det ON gti_notifhor_det.notifnro = gti_notifhor.notifnro"
'            StrSql = StrSql & " AND gti_notifhor.fechadesde <= " & ConvFecha(fechahasta) & " and"
'            StrSql = StrSql & " (gti_notifhor.fechahasta is NULL or gti_notifhor.fechahasta >= " & ConvFecha(fechadesde) & ")"
'            StrSql = StrSql & " AND gti_notifhor.notificado = 0"
'            StrSql = StrSql & " AND gti_notifhor.ternro = " & Ternro & " AND gti_notifhor.empleg = " & emplegajo
'            StrSql = StrSql & " ORDER BY gti_notifhor_det.wd ASC"


            Call vaciarArreglos

            Flog.writeline Espacios(Tabulador * 2) & "Procesando cambios de turno para el empleado"
                
            StrSql = " SELECT * FROM WC_MOV_HORARIOS LEFT JOIN gti_tiptar on gti_tiptar.tptrnro = WC_MOV_HORARIOS.tiptarnro"
            StrSql = StrSql & " WHERE fecdesde = " & ConvFecha(fechadesde) & " AND fechasta "
            StrSql = StrSql & " = " & ConvFecha(fechahasta) & " AND ternro = " & Ternro & " ORDER BY fechor ASC "

            If objRs2.State = adStateOpen Then objRs2.Close
            OpenRecordset StrSql, objRs2
            If Not objRs2.EOF Then 'hay informacion para cargar en los arreglos.
                  
                Flog.writeline Espacios(Tabulador * 2) & "Recuperando datos de la tarjeta del empleado"
                 
                nrotarjeta = objRs2!idtarjeta
                
                If objRs2!tptrdes <> "" Then
                    tipotarjeta = objRs2!tptrdes
                Else
                    Flog.writeline Espacios(Tabulador * 2) & "No se encontro el tipo de tarjeta asociado al empleado."
                End If
                                 
                 puntero = fechadesde
                
                 dias = 14 'dias que faltan comprobar
                
                 Do While (Not objRs2.EOF) And (dias > 0)
                 If ConvFecha(puntero) = ConvFecha(objRs2!fechor) Then
                     Dia.Fecha = puntero
                     Dia.eslaborable = True
                     Dia.descanso = objRs2!desmin
                     Dia.horafin = objRs2!horfin
                     Dia.HoraIni = objRs2!horini
                     If dias > 7 Then 'semana 1
                         semana1(Weekday(puntero)) = Dia
                     Else 'semana 2
                         semana2(Weekday(puntero)) = Dia
                     End If
                     objRs2.MoveNext
                 End If
                 dias = dias - 1
                 puntero = DateAdd("d", 1, puntero)
                 Loop
                  
            End If 'termine de procesar los dias
            
            Call procesarCambios
            
            If hayCambios() And alMenosUnLaborable() Then
            
                fefectiva = DateAdd("d", 7, fechadesde)
                numaccion = "000000"
                          
                Call generarDatosCab(empresarazon, tipotarjeta, nrotarjeta, ccosto, nomemp, puesto, tpago, numaccion, fefectiva)
                repid = getLastIdentity(objConn, "rep_gti_mov_turnos")
                For i = 1 To 7
                    If semana2(i).eslaborable Then
                        Call generarDatosDet(repid, i, semana2(i).HoraIni, semana2(i).horafin, semana2(i).descanso)
                    End If
                Next
                
            Else
                Flog.writeline Espacios(Tabulador * 2) & "No se registraron cambios de turnos para el empleado"
            End If
                      
siguiente:
        
        Progreso = Progreso + IncPorc
            
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & ", bprcempleados = bprcempleados - 1 WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
            
        If SinError Then
             ' borro
             StrSql = "DELETE FROM batch_empleado WHERE ternro = " & Ternro & " AND bpronro = " & NroProceso
             objConnProgreso.Execute StrSql, , adExecuteNoRecords
        Else
             StrSql = "UPDATE batch_empleado SET estado = 'Error' WHERE ternro = " & Ternro & " AND bpronro = " & NroProceso
             objConnProgreso.Execute StrSql, , adExecuteNoRecords
        End If
        
        objRsEmpleado.MoveNext
        Loop
        
    Else
        Exit Sub
    End If
   
    If Not HuboErrores Then
        'Actualizo el estado del proceso
        StrSql = "UPDATE batch_proceso SET bprcprogreso =100 , bprcempleados ='0', bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        StrSql = "UPDATE batch_proceso SET bprcprogreso =100 , bprcempleados ='0', bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    Flog.writeline "TERMINO EL PROCESO CORRECTAMENTE"
    Flog.writeline
    Flog.writeline "Fin : " & Now
    Flog.Close
    objConn.Close
    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now

End Sub

Function okConfRep()

Dim ok As Boolean

    If tipodepago = -1 Then
        ok = False
    Else
        ok = True
    End If
    
    okConfRep = ok
        
End Function
'--------------------------------------------------------------------
' Se encarga de generar los datos
'--------------------------------------------------------------------
Sub generarDatosCab(ByVal emprazon, ByVal tipotarjeta, ByVal nrotarjeta, ByVal ccosto, ByVal nomemp, ByVal puesto, ByVal tpago, _
ByVal numaccion, ByVal fefectiva)


Dim StrSql As String
Dim i As Integer

    '------------------------------------------------------------------
    'Armo la SQL para guardar los datos
    '------------------------------------------------------------------
    StrSql = " INSERT INTO rep_gti_mov_turnos "
    StrSql = StrSql & " ( bpronro, empleg, fdesde, fhasta, empresa, usuario, fcreacion, emprazon, tiptarj, nrotarj,"
    StrSql = StrSql & " ccosto, nomemp, puesto, tpago, numaccion, fefectiva, notif ) VALUES ( "
    StrSql = StrSql & NroProceso
    StrSql = StrSql & "," & emplegajo
    StrSql = StrSql & "," & ConvFecha(fechadesde)
    StrSql = StrSql & "," & ConvFecha(fechahasta)
    StrSql = StrSql & "," & empresa
    StrSql = StrSql & ",'" & usuario & "'"
    StrSql = StrSql & "," & ConvFecha(fechabatch)
    StrSql = StrSql & ",'" & emprazon & "'"
    StrSql = StrSql & ",'" & tipotarjeta & "'"
    StrSql = StrSql & ",'" & nrotarjeta & "'"
    StrSql = StrSql & ",'" & ccosto & "'"
    StrSql = StrSql & ",'" & nomemp & "'"
    StrSql = StrSql & ",'" & puesto & "'"
    StrSql = StrSql & ",'" & tpago & "'"
    StrSql = StrSql & ",'" & numaccion & "'"
    StrSql = StrSql & "," & ConvFecha(fefectiva) & ",0"
    StrSql = StrSql & ")"
        
    '------------------------------------------------------------------
    'Guardo los datos en la BD
    '------------------------------------------------------------------
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline Espacios(Tabulador) & "Se procesó correctamente el empleado " & emplegajo
    Flog.writeline
    Exit Sub

MError:
    Flog.writeline "Error en empleado: " & emplegajo & " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub

End Sub

Sub generarDatosDet(ByVal repid, ByVal wd, ByVal wddesde, ByVal wdhasta, ByVal wdpausa)

Dim StrSql As String

    '------------------------------------------------------------------
    'Armo la SQL para guardar los datos
    '------------------------------------------------------------------
    StrSql = " INSERT INTO rep_gti_mov_turnos_det "
    StrSql = StrSql & " ( repid, empleado, wd, wddesde, wdhasta, wdpausa ) VALUES ("
    StrSql = StrSql & repid
    StrSql = StrSql & "," & emplegajo
    StrSql = StrSql & "," & wd
    StrSql = StrSql & ",'" & wddesde & "'"
    StrSql = StrSql & ",'" & wdhasta & "'"
    StrSql = StrSql & "," & wdpausa
    StrSql = StrSql & ")"
    
    '------------------------------------------------------------------
    'Guardo los datos en la BD
    '------------------------------------------------------------------
    objConn.Execute StrSql, , adExecuteNoRecords
    Exit Sub

MError:
    Flog.writeline "Error en empleado: " & emplegajo & " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub
    
End Sub

Sub vaciarArreglos()

Dim i
Dim Dia As TDia

Dia.eslaborable = False

For i = 1 To 7
    semana1(i) = Dia
    semana2(i) = Dia
    cambios(i) = False
Next


End Sub

Sub procesarCambios()

Dim i


For i = 1 To 7
    'caso en el que algun dia era y dejo de ser laborable, o no era y paso a ser laborable.
    If semana1(i).eslaborable <> semana2(i).eslaborable Then
        cambios(i) = True
    Else 'o era y sigue siendo no laborable, o hay un cambio en el horario para ese dia.
    
        'tengo que analizar si efectivamente hay un cambio
        If (semana1(i).eslaborable = True) And (semana2(i).eslaborable = True) Then
            If (semana1(i).descanso <> semana2(i).descanso) Or (semana1(i).HoraIni <> semana2(i).HoraIni) Or (semana1(i).horafin <> semana2(i).horafin) Then
                cambios(i) = True
            Else
                cambios(i) = False
            End If
        Else ' no era laborable y sigue siendo no laborable
            cambios(i) = False
        End If
    End If
Next
End Sub

Function hayCambios()

Dim i
Dim hay

hay = False
i = 1

 Do While (Not hay) And (i <= 7)
    hay = cambios(i)
    i = i + 1
 Loop

hayCambios = hay

End Function

Function alMenosUnLaborable()

Dim i
Dim hay

hay = False
i = 1

 Do While (Not hay) And (i <= 7)
    hay = semana2(i).eslaborable
    i = i + 1
 Loop

alMenosUnLaborable = hay

End Function
