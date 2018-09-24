Attribute VB_Name = "MdlDepuracionNov"
Option Explicit
'Version: 1.00
'
'
'Const Version = 1.00
'Const FechaVersion = "13/09/2005"

'Const Version = 1.01
'Const FechaVersion = "25/01/2006"   'Version inicial

'Const Version = 1.02
'Const FechaVersion = "26/07/2007"   '26/07/2007 - G. Bauer - Se modifico la decalracion de la variable de Integer a Long

'Const Version = 1.03
'Const FechaVersion = "12/02/2009"   '12/02/2009 - Martin Ferraro - Llamar a la depuracion con dos niveles

'Const Version = 1.04
'Const FechaVersion = "13/02/2009"   '13/02/2009 - Martin Ferraro - Depurar nivel

'Const Version = 1.05
'Const FechaVersion = "17/02/2009"   '17/02/2009 - Martin Ferraro - Correccion error version ant

'Const Version = 1.06
'Const FechaVersion = "04/03/2009"   '04/03/2009 - Martin Ferraro - Nuevas Correcciones

'Global Const Version = 1.07
'Global Const FechaVersion = "12/08/2009"   'Encriptacion de string connection
'Global Const UltimaModificacion = "Manuel Lopez"
'Global Const UltimaModificacion1 = "Encriptacion de string connection"

'Global Const Version = 1.08
'Global Const FechaVersion = "09/11/2009"   'novedades ajustes con multiple empresas
'Global Const UltimaModificacion = "Martin Ferraro"
'Global Const UltimaModificacion1 = ""

'Global Const Version = 1.09
'Global Const FechaVersion = "23/04/2010"
'Global Const UltimaModificacion = "Margiotta, Emanuel"
'Global Const UltimaModificacion1 = "Se corrigio el Procesamiento de Novedades de Ajuste porque no Distinguia si se " & _
'                                    "Filtraba por concepto o por Empleado"
            
            
'Global Const Version = "1.10"
'Global Const FechaVersion = "23/06/2011"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = "Eliminacion de firmas cuando las novedades tienen autorizacion "


'Global Const Version = "1.11"
'Global Const FechaVersion = "21/10/2011"
'Global Const UltimaModificacion = "Margiotta, Emanuel"
'Global Const UltimaModificacion1 = "Se agregó la funcion CreaVistaEmpleadoProceso para trabajar con la vista del usuario que dispara el proceso"

'Global Const Version = "1.12"
'Global Const FechaVersion = "12/06/2014"
'Global Const UltimaModificacion = "Carmen Quintero - CAS-24538 - CCU - BUG EN EL SERVIDOR DE APLICACIONES -"
'Global Const UltimaModificacion1 = " Se agregó en el update final, el progreso en 100 cuando no hay errores. "

'Global Const Version = "1.13"
'Global Const FechaVersion = "07/08/2014"
'Global Const UltimaModificacion = "Fernandez, Matias & Miriam Ruiz- CAS-24538 - CCU - ERROR EN DEPURACION DE NOVEDADES "
'Global Const UltimaModificacion1 = "Si la vista v_empleadoproc ya existe, no se crea."
                 
 
'Global Const Version = "1.14"
'Global Const FechaVersion = "11/08/2014"
'Global Const UltimaModificacion = "Fernandez, Matias & Miriam Ruiz- CAS-24538 - CCU - ERROR EN DEPURACION DE NOVEDADES "
'Global Const UltimaModificacion1 = "Si la vista v_empleadoproc ya existe se borra y se vuelve a crear nuevamente."

'Global Const Version = "1.15"
'Global Const FechaVersion = "12/08/2014"
'Global Const UltimaModificacion = "Fernandez, Matias - CAS-24538 - CCU - ERROR EN DEPURACION DE NOVEDADES "
'Global Const UltimaModificacion1 = "pasa a minusculas suser_sname() en data acces"

Global Const Version = "1.16"
Global Const FechaVersion = "29/08/2014"
Global Const UltimaModificacion = "Quintero, Carmen - CAS-26949 - Arlei - Bug porcentaje de avance proceso Depuración novedades "
Global Const UltimaModificacion1 = "Se modificó la manera de actualizar el progreso"



'-------------------------------------------------------------------------------------------------------------


Private Type ParesConceptoParametro
    ConcNro As Long
    tpanro As Long
End Type

Private Type ParesConcepto
    ConcNro As Long
End Type


Global idUser As String
Global Fecha As Date
Global hora As String

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo As Date
Global StrSql2 As String
Global Pares() As ParesConceptoParametro
Global ConNovAju() As ParesConcepto
Global Porcentaje As Integer
Global Cantidad_Depuradas As Long

Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Proceso.
' Autor      : FGZ
' Fecha      : 18/11/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim ArrParametros
 
On Error GoTo MCE

    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
    If UBound(ArrParametros) > 1 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProcesoBatch = ArrParametros(0)
            Etiqueta = ArrParametros(1)
            EncriptStrconexion = CBool(ArrParametros(2))
            c_seed = ArrParametros(2)
        Else
            Exit Sub
        End If
    Else
        If UBound(ArrParametros) > 0 Then
            If IsNumeric(ArrParametros(0)) Then
                NroProcesoBatch = ArrParametros(0)
                Etiqueta = ArrParametros(1)
            Else
                Exit Sub
            End If
        Else
            If IsNumeric(strCmdLine) Then
                NroProcesoBatch = strCmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    
    'Abro la conexion
On Error Resume Next
OpenConnection strconexion, objConn
If Err.Number <> 0 Or Error_Encrypt Then
 Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
Exit Sub
End If
    
OpenConnection strconexion, objconnProgreso
If Err.Number <> 0 Or Error_Encrypt Then
 Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
Exit Sub
End If
    
    Nombre_Arch = PathFLog & "Depuracion_Novedades" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Modificacion = " & UltimaModificacion1
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline

    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
      Flog.writeline "Cambio el estado del proceso:" & StrSql
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 63 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
      Flog.writeline "Obtengo los datos del proceso:" & StrSql
    TiempoInicialProceso = GetTickCount
    Progreso = 0
    
    If Not rs_batch_proceso.EOF Then
        idUser = rs_batch_proceso!idUser
        Fecha = rs_batch_proceso!bprcfecha
        hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
        Call LevantarParamteros(NroProcesoBatch, bprcparam, idUser)
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.writeline Espacios(Tabulador * 0) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.Close
    objconnProgreso.Close
    objConn.Close
    
    Set objConn = Nothing
    Set objconnProgreso = Nothing
    
    Exit Sub

MCE:
    MyRollbackTrans
    Flog.writeline "Error:" & Err.Description
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
    
    objconnProgreso.Close
    objConn.Close
    
    Set objConn = Nothing
    Set objconnProgreso = Nothing
    
End Sub


Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal parametros As String, ByVal idUser As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : FGZ
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim Separador As String

Dim TipoNovedad As Integer
Dim Todas As Boolean
Dim Filtra_ConceptoParametro As Boolean
Dim Filtro As String
Dim Filtra_Empleados As Boolean
Dim ConVigencia As Boolean
Dim SinVigencia As Boolean
Dim TipoVigencia As Integer
Dim Fecha_Tope As Date
Dim aux As String
Dim Empresa As String

'Son 10 parametros fijos
'Orden
' 1 -   Tipo Novedad:
'                       1 (Ajuste)
'                       2 (Individual)
' 2 -   Todas:
'                       -1 ( para no de ajuste significa todas las novaju y
'                            para nov individuales significa todas las depurables)
'                        0 ( Significa que hay un filtro)
' 3 -   Filtra_Concepto_Parametro:
'                       -1 ( significa que sigue una lista de pares "concnro-parnro" separados por ;)
'                        0 ( Significa que no se filtra por concepto y parametro)
' 4 -   Filtro: Lista de pares "concnro-parnro" separados por , encerrados entre parentesis. Puede venir vacia, es decir ().
' 5 -   Filtra_Empleados:
'                       -1 ( significa que filtro por empleados. La lista esta en Batch_empleado)
'                        0 ( significa que NO filtro por empleados.)
' 6 -   SinVigencia:
'                       -1 ( significa que depuro las novedades que no tienen vigencia)
'                        0 ( significa que NO depuro las novedades que no tienen vigencia)
' 7 -   ConVigencia:
'                       -1 ( significa que depuro las novedades que tienen vigencia)
'                        0 ( significa que NO depuro las novedades que tienen vigencia)
' 8 -   Tipo:
'                       0 (sin vigencia. cuando la opcion anterior es 0)
'                       1 (Todas la novedades que tengan vigencia)
'                       2 (Todas la novedades que tengan vigencia que sean DEPURABLES)
'                       3 (Todas la novedades que tengan vigencia que sean NO DEPURABLES)
' 9 -   Fecha:          (o es 0 o trae una fecha. cuando la opcion anterior es 2 trae una fecha sino trae 0)
'10 -   Empresa         (lista de estrnro separados por ,)

Separador = "@"
' Levanto cada parametro por separado
Flog.writeline "Parametros:"
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, parametros, Separador) - 1
        TipoNovedad = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
         Flog.writeline "TipoNovedad:" & TipoNovedad
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Todas = Mid(parametros, pos1, pos2 - pos1 + 1)
        Flog.writeline "Todas:" & Todas
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Filtra_ConceptoParametro = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        Flog.writeline "Filtra_ConceptoParametro:" & Filtra_ConceptoParametro
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Filtro = Mid(parametros, pos1, pos2 - pos1 + 1)
         Flog.writeline "Filtro:" & Filtro
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Filtra_Empleados = Mid(parametros, pos1, pos2 - pos1 + 1)
           Flog.writeline "Filtra_Empleados:" & Filtra_Empleados
        'EAM- Si selecciono todos los empleado crea la vista temporal
        If Filtra_Empleados = 0 Then
            Call CreaVistaEmpleadoProceso("V_EMPLEADO", idUser)
        End If
           Flog.writeline "Se creo la vista empleados"
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        SinVigencia = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
           Flog.writeline "SinVigencia:" & SinVigencia
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        ConVigencia = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
           Flog.writeline "ConVigencia:" & ConVigencia
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        TipoVigencia = Mid(parametros, pos1, pos2 - pos1 + 1)
         Flog.writeline "TipoVigencia:" & TipoVigencia
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        aux = Mid(parametros, pos1, pos2 - pos1 + 1)
        If TipoVigencia <> 0 Then
            Fecha_Tope = CDate(aux)
        End If
        
        pos1 = pos2 + 2
        pos2 = Len(parametros)
        Empresa = Mid(parametros, pos1, pos2 - pos1 + 1)
        
    End If
End If

''Segun el tipo llamo al depurador correcto
'Select Case TipoNovedad
'Case 1:
'    Call Depurar_Novaju(bpronro, Todas, Filtra_Empleados, ConVigencia, TipoVigencia, Fecha_Tope, Empresa)
'Case 2:
'    Call Depurar_Novemp(bpronro, Todas, Filtra_ConceptoParametro, Filtro, Filtra_Empleados, ConVigencia, TipoVigencia, Fecha_Tope, Empresa)
'Case Else
'
'End Select

'FGZ - 30/12/2004
If TipoNovedad = 0 Then
    'Todas
    Porcentaje = 50
Else
    'alguna de las dos
    Porcentaje = 100
End If
If TipoNovedad = 1 Or TipoNovedad = 0 Then  'De ajuste o Todas
    Call Depurar_Novaju(bpronro, Todas, Filtra_ConceptoParametro, Filtro, Filtra_Empleados, SinVigencia, ConVigencia, TipoVigencia, Fecha_Tope, Empresa)
End If
If TipoNovedad = 2 Or TipoNovedad = 0 Then  'Individuales o todas
    Call Depurar_Novemp(bpronro, Todas, Filtra_ConceptoParametro, Filtro, Filtra_Empleados, SinVigencia, ConVigencia, TipoVigencia, Fecha_Tope, Empresa, 0)
    Call Depurar_Novemp(bpronro, Todas, Filtra_ConceptoParametro, Filtro, Filtra_Empleados, SinVigencia, ConVigencia, TipoVigencia, Fecha_Tope, Empresa, 1)
    Call Depurar_Novemp(bpronro, Todas, Filtra_ConceptoParametro, Filtro, Filtra_Empleados, SinVigencia, ConVigencia, TipoVigencia, Fecha_Tope, Empresa, 2)
End If

End Sub


Public Sub Depurar_Novaju(ByVal bpronro As Long, ByVal Todas As Boolean, ByVal Filtra_Concepto As Boolean, ByVal Filtro As String, ByVal Filtra_Empleados As Boolean, ByVal SinVigencia As Boolean, ByVal ConVigencia As Boolean, ByVal TipoVigencia As Integer, ByVal Fecha_Tope As Date, ByVal Empresa As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de depuracion de novedades de ajuste (novaju)
' Autor      : FGZ
' Fecha      : 18/11/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset
Dim CantCiclos As Integer
Dim Cantidad_Depuradas As Long
Dim MaxConc As Integer
Dim I As Integer
Dim cystipnro As Long

On Error GoTo CE

Cantidad_Depuradas = 0

MyBeginTrans
    If Not Todas Then
        Flog.writeline Espacios(Tabulador * 0) & "Comienzo de la Depuración de las novedades de Ajuste."
        
        If Filtra_Concepto Then
            Call SepararConcepto(Filtro, MaxConc)
        End If
        
        StrSql = "SELECT novaju.nanro, novaju.empleado, novaju.nahasta, novaju.navigencia, conccod "
        StrSql = StrSql & " FROM novaju "
        StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = novaju.concnro"
        
        If Filtra_Concepto And Filtra_Empleados Then
            StrSql = StrSql & " INNER JOIN empleado ON novaju.empleado = empleado.ternro "
            StrSql = StrSql & " INNER JOIN batch_empleado ON empleado.ternro = batch_empleado.ternro"
            StrSql = StrSql & " INNER JOIN his_estructura Empresa ON Empresa.ternro = empleado.ternro "
            StrSql = StrSql & " WHERE Empresa.estrnro IN (" & Empresa & ") AND batch_empleado.bpronro = " & bpronro
        Else
            If Filtra_Concepto Then
                StrSql = StrSql & " INNER JOIN empleado ON novaju.empleado = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
                StrSql = StrSql & " WHERE his_estructura.estrnro IN (" & Empresa & ") "
            End If
            
            If Filtra_Empleados Then
                StrSql = StrSql & " INNER JOIN batch_empleado ON novaju.empleado = batch_empleado.ternro"
                StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = batch_empleado.ternro AND his_estructura.estrnro IN (" & Empresa & ") "
                StrSql = StrSql & " WHERE batch_empleado.bpronro = " & bpronro
            End If
        End If
                
        
        For I = 1 To MaxConc
            If I = 1 Then
                StrSql = StrSql & " AND ( (novaju.concnro = " & ConNovAju(I).ConcNro & ")"
            Else
                StrSql = StrSql & " OR (novaju.concnro = " & ConNovAju(I).ConcNro & ")"
            End If
            If I = MaxConc Then
                StrSql = StrSql & " )"
            End If
        Next I
        OpenRecordset StrSql, rs
               Flog.writeline Espacios(Tabulador * 0) & "novaju-no todas:" & StrSql
        'Setea el incremento del progreso
        CantCiclos = rs.RecordCount
        If CantCiclos = 0 Then
            CantCiclos = 1
        End If
        IncPorc = (Porcentaje / CantCiclos)
        
        Do Until rs.EOF
            If SinVigencia And Not CBool(rs!navigencia) Then
                ' Si no tiene vigencia y la depuracion es sin vigencia, la elimino
                Flog.writeline Espacios(Tabulador * 0) & " La Novedad con concepto " & rs!ConcCod & " para el empleado: " & rs!Empleado & " ,no tiene vigencia y la depuración es sin vigencia por lo tanto se eliminará."
                
                StrSql = "DELETE FROM novaju "
                StrSql = StrSql & " WHERE nanro = " & rs!nanro
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Flog.writeline Espacios(Tabulador * 0) & " La Novedad con concepto " & rs!ConcCod & " para el empleado: " & rs!Empleado & " ha sido depurada."
                
                'FGZ - 22/06/2011 -----------------------
                'Agrego la eliminacion de las firmas
                cystipnro = 20
                StrSql = "DELETE cysfirmas "
                StrSql = StrSql & " where cystipnro = " & cystipnro
                StrSql = StrSql & " and cysfircodext = '" & rs!nanro & "' "
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Firmas eliminadas"
                'FGZ - 22/06/2011 -----------------------
                
                Cantidad_Depuradas = Cantidad_Depuradas + 1
            End If
            'si Tiene fecha de vigencia y la depuracion es con vigencia
            Flog.writeline Espacios(Tabulador * 0) & " Verifica si la Novedad con concepto " & rs!ConcCod & " para el empleado: " & rs!Empleado & " tiene fecha de vigencia y la depuracion es con vigencia."
            
            If CBool(rs!navigencia) And ConVigencia Then
                If DateDiff("d", rs!nahasta, CDate(Fecha_Tope)) >= 0 Then
                    ' Si tiene una fecha de vigencia anterior a la fecha tope, elimino la novedad
                    Flog.writeline Espacios(Tabulador * 0) & " La Novedad con concepto " & rs!ConcCod & " para el empleado: " & rs!Empleado & " tiene una fecha de vigencia anterior a la fecha tope, se elimina la novedad."
                    
                    StrSql = "DELETE FROM novaju "
                    StrSql = StrSql & " WHERE nanro = " & rs!nanro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    Flog.writeline Espacios(Tabulador * 0) & " La Novedad con concepto " & rs!ConcCod & " para el empleado: " & rs!Empleado & " ha sido depurada."
                    
                    'FGZ - 22/06/2011 -----------------------
                    'Agrego la eliminacion de las firmas
                    cystipnro = 20
                    StrSql = "DELETE cysfirmas "
                    StrSql = StrSql & " where cystipnro = " & cystipnro
                    StrSql = StrSql & " and cysfircodext = '" & rs!nanro & "' "
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline Espacios(Tabulador * 1) & "Firmas eliminadas"
                    'FGZ - 22/06/2011 -----------------------
                    
                    Cantidad_Depuradas = Cantidad_Depuradas + 1
                End If
            End If
            
            Progreso = Progreso + IncPorc
            Call ActualizarProgreso(bpronro, Progreso)
            rs.MoveNext
        Loop
    Else 'Todas las novedades
        Flog.writeline Espacios(Tabulador * 0) & "Comienzo de la Depuración de todas las novedades."
        
        StrSql = "SELECT novaju.nanro, novaju.empleado, novaju.nahasta, novaju.navigencia, v_empleadoproc.empleg, conccod "
        StrSql = StrSql & " FROM novaju "
        StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = novaju.concnro"
        StrSql = StrSql & " INNER JOIN v_empleadoproc ON novaju.empleado = v_empleadoproc.ternro"
        StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = v_empleadoproc.ternro AND his_estructura.estrnro IN (" & Empresa & ")"
        OpenRecordset StrSql, rs
           Flog.writeline Espacios(Tabulador * 0) & "todas las novedades:" & StrSql
        'Setea el incremento del progreso
        CantCiclos = rs.RecordCount
        If CantCiclos = 0 Then
            CantCiclos = 1
        End If
        IncPorc = (Porcentaje / CantCiclos)
        
        Do Until rs.EOF
            If SinVigencia And Not CBool(rs!navigencia) Then
                Flog.writeline Espacios(Tabulador * 0) & " La Novedad con concepto " & rs!ConcCod & " para el empleado: " & rs!Empleado & " ,no tiene vigencia y la depuración es sin vigencia por lo tanto se eliminará."
                
                StrSql = "DELETE FROM novaju "
                StrSql = StrSql & " WHERE nanro = " & rs!nanro
                objConn.Execute StrSql, , adExecuteNoRecords
                
                'FGZ - 22/06/2011 -----------------------
                'Agrego la eliminacion de las firmas
                cystipnro = 20
                StrSql = "DELETE cysfirmas "
                StrSql = StrSql & " where cystipnro = " & cystipnro
                StrSql = StrSql & " and cysfircodext = '" & rs!nanro & "' "
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Firmas eliminadas"
                'FGZ - 22/06/2011 -----------------------
                
                
                Cantidad_Depuradas = Cantidad_Depuradas + 1
            End If
            ' Tiene fecha de vigencia
            Flog.writeline Espacios(Tabulador * 0) & " Verifica si la Novedad con concepto " & rs!ConcCod & " para el empleado: " & rs!Empleado & " tiene fecha de vigencia y la depuracion es con vigencia."
            
            If CBool(rs!navigencia) And ConVigencia Then
                If DateDiff("d", rs!nahasta, CDate(Fecha_Tope)) >= 0 Then
                    ' Si tiene una fecha de vigencia anterior a la fecha tope, elimino la novedad
                    Flog.writeline Espacios(Tabulador * 0) & " La Novedad para el empleado: " & rs!Empleado & " tiene una fecha de vigencia anterior a la fecha tope, se elimina la novedad."
                    
                    StrSql = "DELETE FROM novaju "
                    StrSql = StrSql & " WHERE nanro = " & rs!nanro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    Flog.writeline Espacios(Tabulador * 0) & " La Novedad con concepto " & rs!ConcCod & " para el empleado: " & rs!Empleado & " ha sido depurada."
                    
                    'FGZ - 22/06/2011 -----------------------
                    'Agrego la eliminacion de las firmas
                    cystipnro = 20
                    StrSql = "DELETE cysfirmas "
                    StrSql = StrSql & " where cystipnro = " & cystipnro
                    StrSql = StrSql & " and cysfircodext = '" & rs!nanro & "' "
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline Espacios(Tabulador * 1) & "Firmas eliminadas"
                    'FGZ - 22/06/2011 -----------------------
                    
                    
                    Cantidad_Depuradas = Cantidad_Depuradas + 1
                End If
            End If
            
            Progreso = Progreso + IncPorc
            Call ActualizarProgreso(bpronro, Progreso)
            rs.MoveNext
        Loop
    End If
    If rs.State = adStateOpen Then rs.Close
'Fin de la transaccion
MyCommitTrans
If Cantidad_Depuradas > 0 Then
    Flog.writeline Espacios(Tabulador * 1) & Cantidad_Depuradas & " Novedades de Ajuste depuradas "
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró ninguna Novedad de Ajuste para depurar "
End If

If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

Exit Sub
CE:
    HuboError = True
    MyRollbackTrans
    Flog.writeline "Error:" & Err.Description
    Flog.writeline "Ultima SQL ejecuatada: " & StrSql

    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub




Public Sub Depurar_Novemp(ByVal bpronro As Long, ByVal Todas As Boolean, ByVal Filtra_ConceptoParametro As Boolean, ByVal Filtro As String, ByVal Filtra_Empleados As Boolean, ByVal SinVigencia As Boolean, ByVal ConVigencia As Boolean, ByVal TipoVigencia As Integer, ByVal Fecha_Tope As Date, ByVal Empresa As String, ByVal Nivel As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de Depuracion de novedades Individuales (novemp)
' Autor      : FGZ
' Fecha      : 19/11/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset
Dim rsAux As New ADODB.Recordset
Dim I As Integer
Dim MaxPares As Integer
'26/07/2007 - G. Bauer - Se modifico la decalracion de la variable de Integer a Long
'Dim CantCiclos As Integer
Dim CantCiclos As Long
Dim cystipnro As Long

On Error GoTo CE

If Nivel = 0 Then
    Cantidad_Depuradas = 0
End If

MyBeginTrans
    If Not Todas Then
        Flog.writeline Espacios(Tabulador * 0) & "Comienzo de la Depuración de las novedades Nivel " & Nivel & " de Empleado"
       'Filtro activo
        If Filtra_ConceptoParametro Then
            Call SepararConcParam(Filtro, MaxPares)
        End If
        StrSql = "SELECT novemp.*, con_for_tpa.*, conccod "
        StrSql = StrSql & " FROM novemp"
        StrSql = StrSql & " INNER JOIN con_for_tpa ON novemp.concnro = con_for_tpa.concnro AND novemp.tpanro = con_for_tpa.tpanro "
        StrSql = StrSql & " AND con_for_tpa.nivel = " & Nivel
        StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = novemp.concnro"
        If Filtra_Empleados Then
            StrSql = StrSql & " INNER JOIN batch_empleado ON novemp.empleado = batch_empleado.ternro AND batch_empleado.bpronro =" & bpronro
            StrSql = StrSql & " INNER JOIN his_estructura empre ON empre.ternro = batch_empleado.ternro "
        Else
            StrSql = StrSql & " INNER JOIN empleado ON novemp.empleado = empleado.ternro"
            StrSql = StrSql & " INNER JOIN his_estructura empre ON empre.ternro = empleado.ternro "
        End If
        If Nivel = 0 Then
            StrSql = StrSql & " INNER JOIN cft_segun ON cft_segun.concnro = con_for_tpa.concnro AND cft_segun.tpanro = con_for_tpa.tpanro "
            StrSql = StrSql & " AND cft_segun.nivel = con_for_tpa.nivel AND cft_segun.origen = novemp.empleado"
        End If
        If Nivel = 1 Then
            StrSql = StrSql & " INNER JOIN cft_segun ON cft_segun.concnro = con_for_tpa.concnro AND cft_segun.tpanro = con_for_tpa.tpanro "
            StrSql = StrSql & " AND cft_segun.nivel = con_for_tpa.nivel"
            StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro = cft_segun.origen"
            StrSql = StrSql & " AND his_estructura.ternro = novemp.empleado"
            StrSql = StrSql & " AND (his_estructura.htetdesde<=" & ConvFecha(Date) & " AND (his_estructura.htethasta is null or his_estructura.htethasta>=" & ConvFecha(Date) & "))"
        End If
        StrSql = StrSql & " WHERE empre.estrnro IN (" & Empresa & ") "
        For I = 1 To MaxPares
            If I = 1 Then
                StrSql = StrSql & " AND ( (novemp.concnro = " & Pares(I).ConcNro & " AND novemp.tpanro = " & Pares(I).tpanro & ") "
            Else
                StrSql = StrSql & " OR (novemp.concnro = " & Pares(I).ConcNro & " AND novemp.tpanro = " & Pares(I).tpanro & ") "
            End If
            If I = MaxPares Then
                StrSql = StrSql & " )"
            End If
        Next I
        OpenRecordset StrSql, rs
        Flog.writeline Espacios(Tabulador * 0) & "Filtro activo:" & StrSql
        
'        SELECT novemp.*, con_for_tpa.*
'        From novemp
'        INNER JOIN con_for_tpa ON novemp.concnro = con_for_tpa.concnro AND novemp.tpanro = con_for_tpa.tpanro
'        INNER JOIN batch_empleado ON novemp.empleado = batch_empleado.ternro AND batch_empleado.bpronro =
'        INNER JOIN his_estructura ON his_estructura.ternro = batch_empleado.ternro and his_estructura.htethasta is null
'        WHERE his_estructura.estrnro IN (1169)
'        AND ( (novemp.concnro = 11 AND novemp.tpanro = 51)  )
        
        'Setea el incremento del progreso
        CantCiclos = rs.RecordCount
        If CantCiclos = 0 Then
            CantCiclos = 1
        End If
        IncPorc = ((Porcentaje / 3) / CantCiclos)
        
        Do Until rs.EOF
            If Not CBool(rs!nevigencia) And CBool(rs!depurable) And SinVigencia Then
                
                Flog.writeline Espacios(Tabulador * 0) & " Se encontro la novedad con concepto " & rs!ConcCod & " y parámetro " & rs!tpanro & " para el tercero: " & rs!Empleado & "."
                
                'Busca excepciones
                If Nivel = 2 Or Nivel = 1 Then
                
                    '-----------------------------------------------------------------------------------------------------------------------------------------
                    'Buscar que no existe una excepecion para la novedad de Nivel 0 por empleado
                    '-----------------------------------------------------------------------------------------------------------------------------------------
                    StrSql = "SELECT novemp.*, con_for_tpa.*, conccod "
                    StrSql = StrSql & " FROM novemp"
                    StrSql = StrSql & " INNER JOIN con_for_tpa ON novemp.concnro = con_for_tpa.concnro AND novemp.tpanro = con_for_tpa.tpanro "
                    StrSql = StrSql & " AND con_for_tpa.nivel = 0"
                    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = novemp.concnro"
                    If Filtra_Empleados Then
                        StrSql = StrSql & " INNER JOIN batch_empleado ON novemp.empleado = batch_empleado.ternro AND batch_empleado.bpronro =" & bpronro
                        StrSql = StrSql & " INNER JOIN his_estructura empre ON empre.ternro = batch_empleado.ternro "
                    Else
                        StrSql = StrSql & " INNER JOIN empleado ON novemp.empleado = empleado.ternro"
                        StrSql = StrSql & " INNER JOIN his_estructura empre ON empre.ternro = empleado.ternro "
                    End If
                    StrSql = StrSql & " INNER JOIN cft_segun ON cft_segun.concnro = con_for_tpa.concnro AND cft_segun.tpanro = con_for_tpa.tpanro "
                    StrSql = StrSql & " AND cft_segun.nivel = con_for_tpa.nivel AND cft_segun.origen = novemp.empleado"
                    StrSql = StrSql & " WHERE empre.estrnro IN (" & Empresa & ") "
                    StrSql = StrSql & " AND novemp.nenro = " & rs!nenro
                    StrSql = StrSql & " AND depurable <> -1 "
                    For I = 1 To MaxPares
                        If I = 1 Then
                            StrSql = StrSql & " AND ( (novemp.concnro = " & Pares(I).ConcNro & " AND novemp.tpanro = " & Pares(I).tpanro & ") "
                        Else
                            StrSql = StrSql & " OR (novemp.concnro = " & Pares(I).ConcNro & " AND novemp.tpanro = " & Pares(I).tpanro & ") "
                        End If
                        If I = MaxPares Then
                            StrSql = StrSql & " )"
                        End If
                    Next I
                    OpenRecordset StrSql, rsAux
                    Flog.writeline Espacios(Tabulador * 0) & "excepciones:" & StrSql
                    If Not rsAux.EOF Then
                        Flog.writeline Espacios(Tabulador * 1) & "La novedad no se depura ya que tiene excepcion por empleado NO Depurable."
                        GoTo SgtNov
                    End If
                    '-----------------------------------------------------------------------------------------------------------------------------------------
                End If
                
                'Busca excepciones
                If Nivel = 2 Then
                
                    '-----------------------------------------------------------------------------------------------------------------------------------------
                    'Buscar que no existe una excepecion para la novedad de Nivel 1 por estructura
                    '-----------------------------------------------------------------------------------------------------------------------------------------
                    StrSql = "SELECT novemp.*, con_for_tpa.*, conccod "
                    StrSql = StrSql & " FROM novemp"
                    StrSql = StrSql & " INNER JOIN con_for_tpa ON novemp.concnro = con_for_tpa.concnro AND novemp.tpanro = con_for_tpa.tpanro "
                    StrSql = StrSql & " AND con_for_tpa.nivel = 1"
                    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = novemp.concnro"
                    If Filtra_Empleados Then
                        StrSql = StrSql & " INNER JOIN batch_empleado ON novemp.empleado = batch_empleado.ternro AND batch_empleado.bpronro =" & bpronro
                        StrSql = StrSql & " INNER JOIN his_estructura empre ON empre.ternro = batch_empleado.ternro "
                    Else
                        StrSql = StrSql & " INNER JOIN empleado ON novemp.empleado = empleado.ternro"
                        StrSql = StrSql & " INNER JOIN his_estructura empre ON empre.ternro = empleado.ternro "
                    End If
                    StrSql = StrSql & " INNER JOIN cft_segun ON cft_segun.concnro = con_for_tpa.concnro AND cft_segun.tpanro = con_for_tpa.tpanro "
                    StrSql = StrSql & " AND cft_segun.nivel = con_for_tpa.nivel"
                    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro = cft_segun.origen"
                    StrSql = StrSql & " AND his_estructura.ternro = novemp.empleado"
                    StrSql = StrSql & " AND (his_estructura.htetdesde<=" & ConvFecha(Date) & " AND (his_estructura.htethasta is null or his_estructura.htethasta>=" & ConvFecha(Date) & "))"
                    StrSql = StrSql & " WHERE empre.estrnro IN (" & Empresa & ") "
                    StrSql = StrSql & " AND novemp.nenro = " & rs!nenro
                    StrSql = StrSql & " AND depurable <> -1 "
                    For I = 1 To MaxPares
                        If I = 1 Then
                            StrSql = StrSql & " AND ( (novemp.concnro = " & Pares(I).ConcNro & " AND novemp.tpanro = " & Pares(I).tpanro & ") "
                        Else
                            StrSql = StrSql & " OR (novemp.concnro = " & Pares(I).ConcNro & " AND novemp.tpanro = " & Pares(I).tpanro & ") "
                        End If
                        If I = MaxPares Then
                            StrSql = StrSql & " )"
                        End If
                    Next I
                    OpenRecordset StrSql, rsAux
                    Flog.writeline Espacios(Tabulador * 0) & "excepciones por estructura:" & StrSql
                    If Not rsAux.EOF Then
                        Flog.writeline Espacios(Tabulador * 1) & "La novedad no se depura ya que tiene excepcion por estructura NO Depurables"
                        GoTo SgtNov
                    End If
                    '-----------------------------------------------------------------------------------------------------------------------------------------
                    
                End If 'Busca excepciones
                
            
                'Inserto la novedad a borrar en la tabla de backup
                Flog.writeline Espacios(Tabulador * 1) & " Inserto la novedad a borrar en la tabla de backup."
                
                StrSql = "INSERT INTO novemp_bkp ("
                StrSql = StrSql & " concnro"
                StrSql = StrSql & ", tpanro"
                StrSql = StrSql & ", empleado"
                StrSql = StrSql & ", nevalor"
                If Not EsNulo(rs!nevigencia) Then
                    StrSql = StrSql & ", nevigencia"
                End If
                If Not EsNulo(rs!nedesde) Then
                    StrSql = StrSql & ", nedesde"
                End If
                If Not EsNulo(rs!nehasta) Then
                    StrSql = StrSql & ", nehasta"
                End If
                If Not EsNulo(rs!neretro) Then
                    StrSql = StrSql & ", neretro"
                End If
                If Not EsNulo(rs!nepliqdesde) Then
                    StrSql = StrSql & ", nepliqdesde"
                End If
                If Not EsNulo(rs!nepliqhasta) Then
                    StrSql = StrSql & ", nepliqhasta"
                End If
                If Not EsNulo(rs!pronro) Then
                    StrSql = StrSql & ", pronro"
                End If
                StrSql = StrSql & ", nenro"
                If Not EsNulo(rs!netexto) Then
                    StrSql = StrSql & ", netexto"
                End If
                StrSql = StrSql & ", nedepurac"
                
                StrSql = StrSql & " )VALUES( "
                
                StrSql = StrSql & rs!ConcNro
                StrSql = StrSql & "," & rs!tpanro
                StrSql = StrSql & "," & rs!Empleado
                StrSql = StrSql & "," & rs!nevalor
                If Not EsNulo(rs!nevigencia) Then
                    StrSql = StrSql & "," & rs!nevigencia
                End If
                If Not EsNulo(rs!nedesde) Then
                    StrSql = StrSql & "," & ConvFecha(rs!nedesde)
                End If
                If Not EsNulo(rs!nehasta) Then
                    StrSql = StrSql & "," & ConvFecha(rs!nehasta)
                End If
                If Not EsNulo(rs!neretro) Then
                    StrSql = StrSql & "," & ConvFecha(rs!neretro)
                End If
                If Not EsNulo(rs!nepliqdesde) Then
                    StrSql = StrSql & "," & rs!nepliqdesde
                End If
                If Not EsNulo(rs!nepliqhasta) Then
                    StrSql = StrSql & "," & rs!nepliqhasta
                End If
                If Not EsNulo(rs!pronro) Then
                    StrSql = StrSql & "," & rs!pronro
                End If
                StrSql = StrSql & "," & rs!nenro
                If Not EsNulo(rs!netexto) Then
                    StrSql = StrSql & ",'" & rs!netexto & "'"
                End If
                StrSql = StrSql & "," & ConvFecha(CDate(Date))
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                ' Si no tiene fecha de vigencia y es depurable, la elimino
                Flog.writeline Espacios(Tabulador * 1) & "La novedad no tiene fecha de vigencia y es depurable, se elimina."
                
                StrSql = "DELETE FROM novemp "
                StrSql = StrSql & " WHERE nenro = " & rs!nenro
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Flog.writeline Espacios(Tabulador * 1) & "Novedad depurada."
                
                'FGZ - 22/06/2011 -----------------------
                'Agrego la eliminacion de las firmas
                cystipnro = 5
                StrSql = "DELETE cysfirmas "
                StrSql = StrSql & " where cystipnro = " & cystipnro
                StrSql = StrSql & " and cysfircodext = '" & rs!nenro & "' "
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Firmas eliminadas"
                'FGZ - 22/06/2011 -----------------------
                
                
                Cantidad_Depuradas = Cantidad_Depuradas + 1
            End If
            
            If CBool(rs!nevigencia) And ConVigencia Then
                    ' Tiene fecha de vigencia y se selecciono una Fecha Tope
                    If ((TipoVigencia = 2 And CBool(rs!depurable)) Or (TipoVigencia = 3 And Not CBool(rs!depurable)) Or (TipoVigencia = 1)) And (DateDiff("d", rs!nehasta, CDate(Fecha_Tope)) >= 0) Then
'                        StrSql = "INSERT INTO novemp_bkp VALUES(" & rs!concnro & "," & rs!tpanro & "," & rs!Empleado & ","
'                        StrSql = StrSql & GetValor(rs!nevalor, "null") & "," & rs!nevigencia & ","
'                        StrSql = StrSql & GetFecha(rs!nedesde) & "," & GetFecha(rs!nehasta) & ","
'                        StrSql = StrSql & GetFecha(rs!neretro) & "," & GetValor(rs!nepliqdesde, "null") & ","
'                        StrSql = StrSql & GetValor(rs!nepliqhasta, "null") & "," & GetValor(rs!pronro, "null") & ","
'                        StrSql = StrSql & rs!nenro & "," & GetString(rs!netexto) & "," & GetFecha(CDate(Date)) & ")"
                        StrSql = "INSERT INTO novemp_bkp ("
                        StrSql = StrSql & " concnro"
                        StrSql = StrSql & ", tpanro"
                        StrSql = StrSql & ", empleado"
                        StrSql = StrSql & ", nevalor"
                        If Not EsNulo(rs!nevigencia) Then
                            StrSql = StrSql & ", nevigencia"
                        End If
                        If Not EsNulo(rs!nedesde) Then
                            StrSql = StrSql & ", nedesde"
                        End If
                        If Not EsNulo(rs!nehasta) Then
                            StrSql = StrSql & ", nehasta"
                        End If
                        If Not EsNulo(rs!neretro) Then
                            StrSql = StrSql & ", neretro"
                        End If
                        If Not EsNulo(rs!nepliqdesde) Then
                            StrSql = StrSql & ", nepliqdesde"
                        End If
                        If Not EsNulo(rs!nepliqhasta) Then
                            StrSql = StrSql & ", nepliqhasta"
                        End If
                        If Not EsNulo(rs!pronro) Then
                            StrSql = StrSql & ", pronro"
                        End If
                        StrSql = StrSql & ", nenro"
                        If Not EsNulo(rs!netexto) Then
                            StrSql = StrSql & ", netexto"
                        End If
                        StrSql = StrSql & ", nedepurac"
                        
                        StrSql = StrSql & " )VALUES( "
                        
                        StrSql = StrSql & rs!ConcNro
                        StrSql = StrSql & "," & rs!tpanro
                        StrSql = StrSql & "," & rs!Empleado
                        StrSql = StrSql & "," & rs!nevalor
                        If Not EsNulo(rs!nevigencia) Then
                            StrSql = StrSql & "," & rs!nevigencia
                        End If
                        If Not EsNulo(rs!nedesde) Then
                            StrSql = StrSql & "," & ConvFecha(rs!nedesde)
                        End If
                        If Not EsNulo(rs!nehasta) Then
                            StrSql = StrSql & "," & ConvFecha(rs!nehasta)
                        End If
                        If Not EsNulo(rs!neretro) Then
                            StrSql = StrSql & "," & ConvFecha(rs!neretro)
                        End If
                        If Not EsNulo(rs!nepliqdesde) Then
                            StrSql = StrSql & "," & rs!nepliqdesde
                        End If
                        If Not EsNulo(rs!nepliqhasta) Then
                            StrSql = StrSql & "," & rs!nepliqhasta
                        End If
                        If Not EsNulo(rs!pronro) Then
                            StrSql = StrSql & "," & rs!pronro
                        End If
                        StrSql = StrSql & "," & rs!nenro
                        If Not EsNulo(rs!netexto) Then
                            StrSql = StrSql & ",'" & rs!netexto & "'"
                        End If
                        StrSql = StrSql & "," & ConvFecha(CDate(Date))
                        StrSql = StrSql & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        ' Si tiene una fecha de vigencia anterior a la fecha tope, elimino la novedad
                        Flog.writeline Espacios(Tabulador * 1) & "La novedad con concepto " & rs!ConcCod & " y parámetro " & rs!tpanro & " para el empleado: " & rs!Empleado & " tiene una fecha de vigencia anterior a la fecha tope, se elimina la novedad."
                        
                        StrSql = "DELETE FROM novemp "
                        StrSql = StrSql & " WHERE nenro = " & rs!nenro
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        Flog.writeline Espacios(Tabulador * 1) & "La novedad con concepto " & rs!ConcCod & " y parámetro " & rs!tpanro & " para el empleado: " & rs!Empleado & " a sido depurada."
                        
                        'FGZ - 22/06/2011 -----------------------
                        'Agrego la eliminacion de las firmas
                        cystipnro = 5
                        StrSql = "DELETE cysfirmas "
                        StrSql = StrSql & " where cystipnro = " & cystipnro
                        StrSql = StrSql & " and cysfircodext = '" & rs!nenro & "' "
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline Espacios(Tabulador * 1) & "Firmas eliminadas"
                        'FGZ - 22/06/2011 -----------------------
                        
                        Cantidad_Depuradas = Cantidad_Depuradas + 1
                    End If
            End If
            
SgtNov:     Progreso = Progreso + IncPorc
            Call ActualizarProgreso(bpronro, Progreso)
            rs.MoveNext
        Loop
        If rs.State = adStateOpen Then rs.Close
    Else 'Todas
        Flog.writeline Espacios(Tabulador * 0) & "Comienzo de la Depuració de todas las novedades."
        
        StrSql = "SELECT novemp.*, con_for_tpa.depurable, conccod "
        StrSql = StrSql & " FROM novemp INNER JOIN con_for_tpa ON novemp.concnro = con_for_tpa.concnro AND novemp.tpanro = con_for_tpa.tpanro "
        StrSql = StrSql & " INNER JOIN v_empleadoproc ON novemp.empleado = v_empleadoproc.ternro "
        StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = novemp.concnro"
        StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = v_empleadoproc.ternro AND his_estructura.estrnro IN (" & Empresa & ") "
        OpenRecordset StrSql, rs
        Flog.writeline Espacios(Tabulador * 0) & "todas:" & StrSql
        'Setea el incremento del progreso
        CantCiclos = rs.RecordCount
        If CantCiclos = 0 Then
            CantCiclos = 1
        End If
        'IncPorc = (Porcentaje / CantCiclos)
        'Agregado 29/08/2014
        IncPorc = ((Porcentaje / 3) / CantCiclos)
        'fin
        Do Until rs.EOF
            ' Recorro las novedades
            If Not CBool(rs!nevigencia) And CBool(rs!depurable) And SinVigencia Then
                'Inserto la novedad a borrar en la tabla de backup
                Flog.writeline Espacios(Tabulador * 0) & " Inserto la novedad con concepto " & rs!ConcCod & " y parámetro " & rs!tpanro & " a borrar en la tabla de backup para el empleado: " & rs!Empleado & "."
                
                StrSql = "INSERT INTO novemp_bkp VALUES(" & rs!ConcNro & "," & rs!tpanro & "," & rs!Empleado & ","
                StrSql = StrSql & GetValor(rs!nevalor, "null") & "," & rs!nevigencia & ","
                StrSql = StrSql & GetFecha(rs!nedesde) & "," & GetFecha(rs!nehasta) & ","
                StrSql = StrSql & GetFecha(rs!neretro) & "," & GetValor(rs!nepliqdesde, "null") & ","
                StrSql = StrSql & GetValor(rs!nepliqhasta, "null") & "," & GetValor(rs!pronro, "null") & ","
                StrSql = StrSql & rs!nenro & "," & GetString(rs!netexto) & "," & GetFecha(CDate(Date)) & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                ' Si no tiene fecha de vigencia y es depurable, la elimino
                Flog.writeline Espacios(Tabulador * 0) & "La novedad con concepto " & rs!ConcCod & " y parámetro " & rs!tpanro & " para el empleado: " & rs!Empleado & " no tiene fecha de vigencia y es depurable, se elimina."
                
                StrSql = "DELETE FROM novemp "
                StrSql = StrSql & " WHERE nenro = " & rs!nenro
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Flog.writeline Espacios(Tabulador * 0) & "La novedad con concepto " & rs!ConcCod & " y parámetro " & rs!tpanro & " para el empleado: " & rs!Empleado & " a sido depurada."
                
                'FGZ - 22/06/2011 -----------------------
                'Agrego la eliminacion de las firmas
                cystipnro = 5
                StrSql = "DELETE cysfirmas "
                StrSql = StrSql & " where cystipnro = " & cystipnro
                StrSql = StrSql & " and cysfircodext = '" & rs!nenro & "' "
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Firmas eliminadas"
                'FGZ - 22/06/2011 -----------------------
                
                
                Cantidad_Depuradas = Cantidad_Depuradas + 1
            End If
            If CBool(rs!nevigencia) And ConVigencia Then
                ' Tiene fecha de vigencia y se selecciono una Fecha Tope
                If ((TipoVigencia = 2 And CBool(rs!depurable)) Or (TipoVigencia = 3 And Not CBool(rs!depurable)) Or (TipoVigencia = 1)) And (DateDiff("d", rs!nehasta, CDate(Fecha_Tope)) >= 0) Then
                    'Inserto la novedad a borrar en la tabla de backup
                    StrSql = "INSERT INTO novemp_bkp VALUES(" & rs!ConcNro & "," & rs!tpanro & "," & rs!Empleado & ","
                    StrSql = StrSql & GetValor(rs!nevalor, "null") & "," & rs!nevigencia & ","
                    StrSql = StrSql & GetFecha(rs!nedesde) & "," & GetFecha(rs!nehasta) & ","
                    StrSql = StrSql & GetFecha(rs!neretro) & "," & GetValor(rs!nepliqdesde, "null") & ","
                    StrSql = StrSql & GetValor(rs!nepliqhasta, "null") & "," & GetValor(rs!pronro, "null") & ","
                    StrSql = StrSql & rs!nenro & "," & GetString(rs!netexto) & "," & GetFecha(CDate(Date)) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords

                    ' Si tiene una fecha de vigencia anterior a la fecha tope, elimino la novedad
                    Flog.writeline Espacios(Tabulador * 0) & "La novedad con concepto " & rs!ConcCod & " y parámetro " & rs!tpanro & " para el empleado: " & rs!Empleado & " tiene una fecha de vigencia anterior a la fecha tope, se elimina la novedad."
                    
                    StrSql = "DELETE FROM novemp "
                    StrSql = StrSql & " WHERE nenro = " & rs!nenro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    Flog.writeline Espacios(Tabulador * 0) & "La novedad con concepto " & rs!ConcCod & " y parámetro " & rs!tpanro & " para el empleado: " & rs!Empleado & " a sido depurada."
                    
                    'FGZ - 22/06/2011 -----------------------
                    'Agrego la eliminacion de las firmas
                    cystipnro = 5
                    StrSql = "DELETE cysfirmas "
                    StrSql = StrSql & " where cystipnro = " & cystipnro
                    StrSql = StrSql & " and cysfircodext = '" & rs!nenro & "' "
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline Espacios(Tabulador * 1) & "Firmas eliminadas"
                    'FGZ - 22/06/2011 -----------------------
                    
                    Cantidad_Depuradas = Cantidad_Depuradas + 1
                End If
            End If
            
            Progreso = Progreso + IncPorc
            Call ActualizarProgreso(bpronro, Progreso)
            rs.MoveNext
        Loop
        rs.Close
    End If
'Fin de la transaccion
MyCommitTrans

If Cantidad_Depuradas > 0 Then
    Flog.writeline Espacios(Tabulador * 1) & Cantidad_Depuradas & "Novedades Individuales depuradas"
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encontraron Novedades Individuales para depurar"
End If

If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

Exit Sub

CE:
    HuboError = True
    MyRollbackTrans
    Flog.writeline "Error:" & Err.Description
    Flog.writeline "Ultima SQL ejecutada: " & StrSql

    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub


Public Sub Depurar_Novaju_old(ByVal bpronro As Long, ByVal Todas As Boolean, ByVal Filtra_Empleados As Boolean, ByVal SinVigencia As Boolean, ByVal ConVigencia As Boolean, ByVal TipoVigencia As Integer, ByVal Fecha_Tope As Date, ByVal Empresa As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de depuracion de novedades de ajuste (novaju)
' Autor      : FGZ
' Fecha      : 18/11/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset
Dim CantCiclos As Integer
Dim Cantidad_Depuradas As Long

On Error GoTo CE

Cantidad_Depuradas = 0
Progreso = 0

MyBeginTrans
    If Not Todas Then
        StrSql = "SELECT novaju.nanro, novaju.empleado, novaju.nahasta, novaju.navigencia "
        StrSql = StrSql & " FROM novaju "
        StrSql = StrSql & " INNER JOIN batch_empleado ON novaju.empleado = batch_empleado.ternro"
        StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = batch_empleado.ternro and his_estructura.htethasta is null AND his_estructura.estrnro IN (" & Empresa & ") "
        StrSql = StrSql & " WHERE batch_empleado.bpronro = " & bpronro
        OpenRecordset StrSql, rs
           Flog.writeline Espacios(Tabulador * 0) & "novedades de ajuste no todas:" & StrSql
        'Setea el incremento del progreso
        CantCiclos = rs.RecordCount
        If CantCiclos = 0 Then
            CantCiclos = 1
        End If
        IncPorc = (100 / CantCiclos)
        
        Do Until rs.EOF
            If Not ConVigencia And Not CBool(rs!navigencia) Then
                ' Si no tiene fecha de vigencia y la depuracion es sin vigencia, la elimino
                StrSql = "DELETE FROM novaju "
                StrSql = StrSql & " WHERE nanro = " & rs!nanro
                objConn.Execute StrSql, , adExecuteNoRecords
                Cantidad_Depuradas = Cantidad_Depuradas + 1
            Else
                'si Tiene fecha de vigencia y la depuracion es con vigencia
                If CBool(rs!navigencia) And ConVigencia Then
                    If DateDiff("d", rs!nahasta, CDate(Fecha_Tope)) >= 0 Then
                        ' Si tiene una fecha de vigencia anterior a la fecha tope, elimino la novedad
                        StrSql = "DELETE FROM novaju "
                        StrSql = StrSql & " WHERE nanro = " & rs!nanro
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Cantidad_Depuradas = Cantidad_Depuradas + 1
                    End If
                End If
            End If
            Progreso = Progreso + IncPorc
            Call ActualizarProgreso(bpronro, Progreso)
            rs.MoveNext
        Loop
    Else 'Todas las novedades
        StrSql = "SELECT novaju.nanro, novaju.empleado, novaju.nahasta, novaju.navigencia, empleado.empleg "
        StrSql = StrSql & " FROM novaju "
        StrSql = StrSql & " INNER JOIN empleado ON novaju.empleado = empleado.ternro"
        StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro and his_estructura.htethasta is null AND his_estructura.estrnro = " & Empresa
        OpenRecordset StrSql, rs
        Flog.writeline Espacios(Tabulador * 0) & "Novedades de ajuste-todas:" & StrSql
        'Setea el incremento del progreso
        CantCiclos = rs.RecordCount
        If CantCiclos = 0 Then
            CantCiclos = 1
        End If
        IncPorc = (100 / CantCiclos)
        
        Do Until rs.EOF
            If Not CBool(rs!navigencia) Then
                ' Si no tiene fecha de vigencia, la elimino
                StrSql = "DELETE FROM novaju "
                StrSql = StrSql & " WHERE nanro = " & rs!nanro
                objConn.Execute StrSql, , adExecuteNoRecords
                Cantidad_Depuradas = Cantidad_Depuradas + 1
            Else
                ' Tiene fecha de vigencia
                If CBool(rs!navigencia) And ConVigencia Then
                    If DateDiff("d", rs!nahasta, CDate(Fecha_Tope)) >= 0 Then
                        ' Si tiene una fecha de vigencia anterior a la fecha tope, elimino la novedad
                        StrSql = "DELETE FROM novaju "
                        StrSql = StrSql & " WHERE nanro = " & rs!nanro
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Cantidad_Depuradas = Cantidad_Depuradas + 1
                    End If
                End If
            End If
            
            Progreso = Progreso + IncPorc
            Call ActualizarProgreso(bpronro, Progreso)
            rs.MoveNext
        Loop
    End If
    If rs.State = adStateOpen Then rs.Close
'Fin de la transaccion
MyCommitTrans
If Cantidad_Depuradas > 0 Then
    Flog.writeline Espacios(Tabulador * 1) & Cantidad_Depuradas & " Novedades de Ajuste depuradas "
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró ninguna Novedad de Ajuste para depurar "
End If

If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

Exit Sub
CE:
    HuboError = True
    MyRollbackTrans
    Flog.writeline "Error:" & Err.Description

    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing

End Sub



Public Sub Depurar_Novemp_old(ByVal bpronro As Long, ByVal Todas As Boolean, ByVal Filtra_ConceptoParametro As Boolean, ByVal Filtro As String, ByVal Filtra_Empleados As Boolean, ByVal SinVigencia As Boolean, ByVal ConVigencia As Boolean, ByVal TipoVigencia As Integer, ByVal Fecha_Tope As Date, ByVal Empresa As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de Depuracion de novedades Individuales (novemp)
' Autor      : FGZ
' Fecha      : 19/11/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset
Dim I As Integer
Dim MaxPares As Integer
Dim CantCiclos As Integer
Dim Cantidad_Depuradas As Long

On Error GoTo CE

Progreso = 0
Cantidad_Depuradas = 0

MyBeginTrans
    If Not Todas Then
       'Filtro activo
        If Filtra_ConceptoParametro Then
            Call SepararConcParam(Filtro, MaxPares)
        End If
        StrSql = "SELECT novemp.*, con_for_tpa.* "
        StrSql = StrSql & " FROM novemp INNER JOIN con_for_tpa ON novemp.concnro = con_for_tpa.concnro AND novemp.tpanro = con_for_tpa.tpanro "
        If Filtra_Empleados Then
            StrSql = StrSql & " INNER JOIN batch_empleado ON novemp.empleado = batch_empleado.ternro AND batch_empleado.bpronro =" & bpronro
            StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = batch_empleado.ternro and his_estructura.htethasta is null "
        Else
            StrSql = StrSql & " INNER JOIN empleado ON novemp.empleado = empleado.ternro"
            StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro and his_estructura.htethasta is null "
        End If
        StrSql = StrSql & " WHERE his_estructura.estrnro IN (" & Empresa & ") "
        For I = 1 To MaxPares
            If I = 1 Then
                StrSql = StrSql & " AND ( (novemp.concnro = " & Pares(I).ConcNro & " AND novemp.tpanro = " & Pares(I).tpanro & ") "
            Else
                StrSql = StrSql & " OR (novemp.concnro = " & Pares(I).ConcNro & " AND novemp.tpanro = " & Pares(I).tpanro & ") "
            End If
            If I = MaxPares Then
                StrSql = StrSql & " )"
            End If
        Next I
        OpenRecordset StrSql, rs
        Flog.writeline Espacios(Tabulador * 0) & "novedades individuales-no todas:" & StrSql
        
'        SELECT novemp.*, con_for_tpa.*
'        From novemp
'        INNER JOIN con_for_tpa ON novemp.concnro = con_for_tpa.concnro AND novemp.tpanro = con_for_tpa.tpanro
'        INNER JOIN batch_empleado ON novemp.empleado = batch_empleado.ternro AND batch_empleado.bpronro =
'        INNER JOIN his_estructura ON his_estructura.ternro = batch_empleado.ternro and his_estructura.htethasta is null
'        WHERE his_estructura.estrnro IN (1169)
'        AND ( (novemp.concnro = 11 AND novemp.tpanro = 51)  )
        
        'Setea el incremento del progreso
        CantCiclos = rs.RecordCount
        If CantCiclos = 0 Then
            CantCiclos = 1
        End If
        IncPorc = (100 / CantCiclos)
        
        Do Until rs.EOF
            If Not CBool(rs!nevigencia) And CBool(rs!depurable) Then
                'Inserto la novedad a borrar en la tabla de backup
                StrSql = "INSERT INTO novemp_bkp ("
                StrSql = StrSql & " concnro"
                StrSql = StrSql & ", tpanro"
                StrSql = StrSql & ", empleado"
                StrSql = StrSql & ", nevalor"
                If Not EsNulo(rs!nevigencia) Then
                    StrSql = StrSql & ", nevigencia"
                End If
                If Not EsNulo(rs!nedesde) Then
                    StrSql = StrSql & ", nedesde"
                End If
                If Not EsNulo(rs!nehasta) Then
                    StrSql = StrSql & ", nehasta"
                End If
                If Not EsNulo(rs!neretro) Then
                    StrSql = StrSql & ", neretro"
                End If
                If Not EsNulo(rs!nepliqdesde) Then
                    StrSql = StrSql & ", nepliqdesde"
                End If
                If Not EsNulo(rs!nepliqhasta) Then
                    StrSql = StrSql & ", nepliqhasta"
                End If
                If Not EsNulo(rs!pronro) Then
                    StrSql = StrSql & ", pronro"
                End If
                StrSql = StrSql & ", nenro"
                If Not EsNulo(rs!netexto) Then
                    StrSql = StrSql & ", netexto"
                End If
                StrSql = StrSql & ", nedepurac"
                
                StrSql = StrSql & " )VALUES( "
                
                StrSql = StrSql & rs!ConcNro
                StrSql = StrSql & "," & rs!tpanro
                StrSql = StrSql & "," & rs!Empleado
                StrSql = StrSql & "," & rs!nevalor
                If Not EsNulo(rs!nevigencia) Then
                    StrSql = StrSql & "," & rs!nevigencia
                End If
                If Not EsNulo(rs!nedesde) Then
                    StrSql = StrSql & "," & ConvFecha(rs!nedesde)
                End If
                If Not EsNulo(rs!nehasta) Then
                    StrSql = StrSql & "," & ConvFecha(rs!nehasta)
                End If
                If Not EsNulo(rs!neretro) Then
                    StrSql = StrSql & "," & ConvFecha(rs!neretro)
                End If
                If Not EsNulo(rs!nepliqdesde) Then
                    StrSql = StrSql & "," & rs!nepliqdesde
                End If
                If Not EsNulo(rs!nepliqhasta) Then
                    StrSql = StrSql & "," & rs!nepliqhasta
                End If
                If Not EsNulo(rs!pronro) Then
                    StrSql = StrSql & "," & rs!pronro
                End If
                StrSql = StrSql & "," & rs!nenro
                If Not EsNulo(rs!netexto) Then
                    StrSql = StrSql & ",'" & rs!netexto & "'"
                End If
                StrSql = StrSql & "," & ConvFecha(CDate(Date))
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                ' Si no tiene fecha de vigencia y es depurable, la elimino
                StrSql = "DELETE FROM novemp "
                StrSql = StrSql & " WHERE nenro = " & rs!nenro
                objConn.Execute StrSql, , adExecuteNoRecords
                Cantidad_Depuradas = Cantidad_Depuradas + 1
            Else
                If CBool(rs!nevigencia) And ConVigencia Then
                    ' Tiene fecha de vigencia y se selecciono una Fecha Tope
                    If ((TipoVigencia = 2 And CBool(rs!depurable)) Or (TipoVigencia = 3 And Not CBool(rs!depurable)) Or (TipoVigencia = 1)) And (DateDiff("d", rs!nehasta, CDate(Fecha_Tope)) >= 0) Then
'                        StrSql = "INSERT INTO novemp_bkp VALUES(" & rs!concnro & "," & rs!tpanro & "," & rs!Empleado & ","
'                        StrSql = StrSql & GetValor(rs!nevalor, "null") & "," & rs!nevigencia & ","
'                        StrSql = StrSql & GetFecha(rs!nedesde) & "," & GetFecha(rs!nehasta) & ","
'                        StrSql = StrSql & GetFecha(rs!neretro) & "," & GetValor(rs!nepliqdesde, "null") & ","
'                        StrSql = StrSql & GetValor(rs!nepliqhasta, "null") & "," & GetValor(rs!pronro, "null") & ","
'                        StrSql = StrSql & rs!nenro & "," & GetString(rs!netexto) & "," & GetFecha(CDate(Date)) & ")"
                        StrSql = "INSERT INTO novemp_bkp ("
                        StrSql = StrSql & " concnro"
                        StrSql = StrSql & ", tpanro"
                        StrSql = StrSql & ", empleado"
                        StrSql = StrSql & ", nevalor"
                        If Not EsNulo(rs!nevigencia) Then
                            StrSql = StrSql & ", nevigencia"
                        End If
                        If Not EsNulo(rs!nedesde) Then
                            StrSql = StrSql & ", nedesde"
                        End If
                        If Not EsNulo(rs!nehasta) Then
                            StrSql = StrSql & ", nehasta"
                        End If
                        If Not EsNulo(rs!neretro) Then
                            StrSql = StrSql & ", neretro"
                        End If
                        If Not EsNulo(rs!nepliqdesde) Then
                            StrSql = StrSql & ", nepliqdesde"
                        End If
                        If Not EsNulo(rs!nepliqhasta) Then
                            StrSql = StrSql & ", nepliqhasta"
                        End If
                        If Not EsNulo(rs!pronro) Then
                            StrSql = StrSql & ", pronro"
                        End If
                        StrSql = StrSql & ", nenro"
                        If Not EsNulo(rs!netexto) Then
                            StrSql = StrSql & ", netexto"
                        End If
                        StrSql = StrSql & ", nedepurac"
                        
                        StrSql = StrSql & " )VALUES( "
                        
                        StrSql = StrSql & rs!ConcNro
                        StrSql = StrSql & "," & rs!tpanro
                        StrSql = StrSql & "," & rs!Empleado
                        StrSql = StrSql & "," & rs!nevalor
                        If Not EsNulo(rs!nevigencia) Then
                            StrSql = StrSql & "," & rs!nevigencia
                        End If
                        If Not EsNulo(rs!nedesde) Then
                            StrSql = StrSql & "," & ConvFecha(rs!nedesde)
                        End If
                        If Not EsNulo(rs!nehasta) Then
                            StrSql = StrSql & "," & ConvFecha(rs!nehasta)
                        End If
                        If Not EsNulo(rs!neretro) Then
                            StrSql = StrSql & "," & ConvFecha(rs!neretro)
                        End If
                        If Not EsNulo(rs!nepliqdesde) Then
                            StrSql = StrSql & "," & rs!nepliqdesde
                        End If
                        If Not EsNulo(rs!nepliqhasta) Then
                            StrSql = StrSql & "," & rs!nepliqhasta
                        End If
                        If Not EsNulo(rs!pronro) Then
                            StrSql = StrSql & "," & rs!pronro
                        End If
                        StrSql = StrSql & "," & rs!nenro
                        If Not EsNulo(rs!netexto) Then
                            StrSql = StrSql & ",'" & rs!netexto & "'"
                        End If
                        StrSql = StrSql & "," & ConvFecha(CDate(Date))
                        StrSql = StrSql & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        ' Si tiene una fecha de vigencia anterior a la fecha tope, elimino la novedad
                        StrSql = "DELETE FROM novemp "
                        StrSql = StrSql & " WHERE nenro = " & rs!nenro
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Cantidad_Depuradas = Cantidad_Depuradas + 1
                    End If
                End If
            End If
            
            Progreso = Progreso + IncPorc
            Call ActualizarProgreso(bpronro, Progreso)
            rs.MoveNext
        Loop
        If rs.State = adStateOpen Then rs.Close
    Else 'Todas
        StrSql = "SELECT novemp.*, con_for_tpa.depurable "
        StrSql = StrSql & " FROM novemp INNER JOIN con_for_tpa ON novemp.concnro = con_for_tpa.concnro AND novemp.tpanro = con_for_tpa.tpanro "
        StrSql = StrSql & " INNER JOIN empleado ON novemp.empleado = empleado.ternro "
        StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro and his_estructura.htethasta is null AND his_estructura.estrnro IN (" & Empresa & ") "
        OpenRecordset StrSql, rs
        Flog.writeline Espacios(Tabulador * 0) & "Novedades indivuduales-todas:" & StrSql
        'Setea el incremento del progreso
        CantCiclos = rs.RecordCount
        If CantCiclos = 0 Then
            CantCiclos = 1
        End If
        IncPorc = (100 / CantCiclos)
        
        Do Until rs.EOF
            ' Recorro las novedades
            If Not CBool(rs!nevigencia) And CBool(rs!depurable) Then
                'Inserto la novedad a borrar en la tabla de backup
                StrSql = "INSERT INTO novemp_bkp VALUES(" & rs!ConcNro & "," & rs!tpanro & "," & rs!Empleado & ","
                StrSql = StrSql & GetValor(rs!nevalor, "null") & "," & rs!nevigencia & ","
                StrSql = StrSql & GetFecha(rs!nedesde) & "," & GetFecha(rs!nehasta) & ","
                StrSql = StrSql & GetFecha(rs!neretro) & "," & GetValor(rs!nepliqdesde, "null") & ","
                StrSql = StrSql & GetValor(rs!nepliqhasta, "null") & "," & GetValor(rs!pronro, "null") & ","
                StrSql = StrSql & rs!nenro & "," & GetString(rs!netexto) & "," & GetFecha(CDate(Date)) & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                ' Si no tiene fecha de vigencia y es depurable, la elimino
                StrSql = "DELETE FROM novemp "
                StrSql = StrSql & " WHERE nenro = " & rs!nenro
                objConn.Execute StrSql, , adExecuteNoRecords
                Cantidad_Depuradas = Cantidad_Depuradas + 1
            Else
                If CBool(rs!nevigencia) And ConVigencia Then
                    ' Tiene fecha de vigencia y se selecciono una Fecha Tope
                    If ((TipoVigencia = 2 And CBool(rs!depurable)) Or (TipoVigencia = 3 And Not CBool(rs!depurable)) Or (TipoVigencia = 1)) And (DateDiff("d", rs!nehasta, CDate(Fecha_Tope)) >= 0) Then
                        'Inserto la novedad a borrar en la tabla de backup
                        StrSql = "INSERT INTO novemp_bkp VALUES(" & rs!ConcNro & "," & rs!tpanro & "," & rs!Empleado & ","
                        StrSql = StrSql & GetValor(rs!nevalor, "null") & "," & rs!nevigencia & ","
                        StrSql = StrSql & GetFecha(rs!nedesde) & "," & GetFecha(rs!nehasta) & ","
                        StrSql = StrSql & GetFecha(rs!neretro) & "," & GetValor(rs!nepliqdesde, "null") & ","
                        StrSql = StrSql & GetValor(rs!nepliqhasta, "null") & "," & GetValor(rs!pronro, "null") & ","
                        StrSql = StrSql & rs!nenro & "," & GetString(rs!netexto) & "," & GetFecha(CDate(Date)) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords

                        ' Si tiene una fecha de vigencia anterior a la fecha tope, elimino la novedad
                        StrSql = "DELETE FROM novemp "
                        StrSql = StrSql & " WHERE nenro = " & rs!nenro
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Cantidad_Depuradas = Cantidad_Depuradas + 1
                    End If
                End If
            End If
            
            Progreso = Progreso + IncPorc
            Call ActualizarProgreso(bpronro, Progreso)
            rs.MoveNext
        Loop
        rs.Close
    End If
'Fin de la transaccion
MyCommitTrans

If Cantidad_Depuradas > 0 Then
    Flog.writeline Espacios(Tabulador * 1) & Cantidad_Depuradas & "Novedades Individuales depuradas"
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encontraron Novedades Individuales para depurar"
End If

If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

Exit Sub
CE:
    HuboError = True
    MyRollbackTrans
    Flog.writeline "Error:" & Err.Description

    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub



Public Sub SepararConcParam(ByVal lista As String, ByRef Cantidad As Integer)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para separar la lista de Concepto-Parametro.
' Autor      : FGZ
' Fecha      : 19/11/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim arr
Dim arr2
Dim I As Integer

arr = Split(lista, ",")
Cantidad = UBound(arr) - 1
ReDim Preserve Pares(Cantidad) As ParesConceptoParametro
If Cantidad >= 0 Then
    For I = 1 To (UBound(arr) - 1)
        arr2 = Split(arr(I), "-")
        Pares(I).ConcNro = arr2(0)
        Pares(I).tpanro = arr2(1)
    Next
Else
    Cantidad = -1
End If
End Sub

Public Sub SepararConcepto(ByVal lista As String, ByRef Cantidad As Integer)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para separar la lista de Concepto.
' Autor      : JMH
' Fecha      : 13/07/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim arr
Dim arr2
Dim I As Integer

arr = Split(lista, ",")
Cantidad = UBound(arr) - 1
ReDim Preserve ConNovAju(Cantidad) As ParesConcepto
If Cantidad >= 0 Then
    For I = 1 To (UBound(arr) - 1)
        arr2 = Split(arr(I), "-")
        ConNovAju(I).ConcNro = arr2(0)
    Next
Else
    Cantidad = -1
End If
End Sub


Public Sub ActualizarProgreso(ByVal NroProceso As Long, ByVal Progreso As Single)
' --------------------------------------------------------------------------------------------
' Descripcion: Actualizo el progreso del Proceso
' Autor      : FGZ
' Fecha      : 19/11/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
End Sub
