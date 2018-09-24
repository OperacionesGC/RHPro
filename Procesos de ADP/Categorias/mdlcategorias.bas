Attribute VB_Name = "mdlCategorias"
Option Explicit

Const Version = "1.01"
Const FechaVersion = "12/03/2013"
'                    Miriam Ruiz - CAS-28350 - Salto Grande - Custom ADP - Custom en Categorias
'                   se cambió la estructura 3 categoria por la 112 categoria2

'Const Version = "1.00"
'Const FechaVersion = "10/02/2013"
'                    Miriam Ruiz - CAS-28350 - Salto Grande - Custom ADP - Custom en Categorias
' ----------------------------------------------------------------------------------------------------------------



'definicion de variables globales de configuracion basica
' Global strformatoFservidor As String
'Global EncriptStrconexion As Boolean
'Global Error_Encrypt As Boolean
'Global c_seed As String
'Global cprnnro As Long  'Guarda el nro de error de la tabla inter_pin
'Global Ya_Encripto As Boolean
'Global TipoBD As String
                        ' DB2 = 1
                        ' Informix = 2
                        ' SQL Server = 3
'Global PathFLog As String

'FGZ - 18/06/2004
Global NumeroSeparadorDecimal As String
Global NumeroSeparadorMiles As String
Global MonedaSeparadorDecimal As String
Global MonedaSeparadorMiles As String
Global Nuevo_NumeroSeparadorDecimal As String
Global Nuevo_NumeroSeparadorMiles As String
Global Nuevo_MonedaSeparadorDecimal As String
Global Nuevo_MonedaSeparadorMiles As String
Global IncPorc As Single
Global IncPorcEmpleado As Single
Global HuboErrores As Boolean
Global EmpleadoSinError As Boolean
Global Progreso As Single
Global ProgresoEmpleado As Single
Global fec_proc As Integer ' 1 - Política Primer Reg.
                           ' 2 - Política Reg. del Turno
                           ' 3 - Política Ultima Reg.

Global OK As Boolean
Global fecha_desde As Date
Global fecha_hasta As Date



'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Sub Main()

Dim Archivo As String
Dim pos As Integer
Dim strcmdLine  As String

'Dim objconnMain As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim NroProceso As Long
Dim NroReporte As Long
Dim StrParametros As String

Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset
Dim myrs As New ADODB.Recordset

Dim PID As String
Dim ArrParametros
Dim sinerror As Boolean

Dim ternro As Long
Dim FDesde As String
Dim FHasta As String
Dim FDesdeOld As String
Dim FHastaOld As String

Dim cantProgreso As Double
Dim totalProgreso As Double
Dim horaUpdateBD As String

Dim tiempoInicial
Dim tiempoActual

Dim I As Long

Dim ListaPar

    strcmdLine = Command()
    
    ArrParametros = Split(strcmdLine, " ", -1)
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
            If IsNumeric(strcmdLine) Then
                NroProceso = strcmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    TiempoInicialProceso = GetTickCount
    'Creo el archivo de texto del desglose
    Archivo = PathFLog & "Categorias-" & CStr(NroProceso) & ".log"

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Archivo, True)

   
    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Inicio :" & Now
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    Flog.writeline strconexion
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If


    On Error GoTo CE


    

    

    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

    'FGZ - 16/01/2015 ---------------------------------------------------
    'Flog.writeline "Levanta Proceso y Setea Parámetros:  " & " " & Now
    Flog.writeline Espacios(Tabulador * 0) & "Levanta Proceso y Setea Parámetros: " & Now
    'FGZ - 16/01/2015 ---------------------------------------------------
    
    'levanto los parametros del proceso
    StrSql = "SELECT bprcfecdesde,bprcfechasta,bprcparam,iduser FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Flog.writeline "No se encontro el proceso " & NroProceso
        Exit Sub
    End If
     
    ternro = 0
    FDesde = Date
    FHasta = Date
    FDesdeOld = Date
    FHastaOld = Date
    If Not EsNulo(rs!bprcparam) Then
        If InStr(1, rs!bprcparam, "@") <> 0 Then
            ListaPar = Split(rs!bprcparam, "@", -1)
               ternro = ListaPar(0)
               FDesde = ListaPar(1)
               FHasta = ListaPar(2)
               FDesdeOld = ListaPar(3)
               FHastaOld = ListaPar(4)
          
        Else
                Flog.writeline "Error en Parametros."
                Exit Sub
        End If
       
    Else
       Flog.writeline "Error en Parametros."
       Exit Sub
    End If
    


    
     If FHastaOld = "" Then
         FHastaOld = "01/01/2199"
    End If
    If FHasta = "" Then
         FHasta = "01/01/2199"
    End If
       
    Call CalculoCat(ternro, FDesde, FHasta)
    If FDesdeOld < FDesde Then
           Call CalculoCat(ternro, FDesdeOld, DateDiff("d", -1, FHasta))
    End If
    
    If InStr(1, FHastaOld, "2199") > 0 Then
       If Not InStr(1, FHasta, "2199") > 0 Then
          'Call CalculoCat(ternro, DateDiff("d", 1, FHasta), FHastaOld)
       End If
    Else
       If Not InStr(1, FHasta, "2199") > 0 And FHastaOld > FHasta Then
            Call CalculoCat(ternro, DateDiff("d", 1, FHasta), FHastaOld)
       End If
    End If
    
    ' poner el bprcestado en procesado
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado', bprcprogreso = 100 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

    ' -----------------------------------------------------------------------------------
    'Copio el proceso en el historico de batch_proceso y lo borro de batch_proceso
        StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProceso
        OpenRecordset StrSql, rs_Batch_Proceso

        StrSql = "INSERT INTO His_Batch_Proceso (bpronro,btprcnro,bprcfecha,iduser"
        StrSqlDatos = rs_Batch_Proceso!bpronro & "," & rs_Batch_Proceso!btprcnro & "," & _
                 ConvFecha(rs_Batch_Proceso!bprcfecha) & ",'" & rs_Batch_Proceso!IdUser & "'"
        
        If Not IsNull(rs_Batch_Proceso!bprchora) Then
            StrSql = StrSql & ",bprchora"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprchora & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcempleados) Then
            StrSql = StrSql & ",bprcempleados"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcempleados & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecdesde) Then
            StrSql = StrSql & ",bprcfecdesde"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecdesde)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfechasta) Then
            StrSql = StrSql & ",bprcfechasta"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfechasta)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcestado) Then
            StrSql = StrSql & ",bprcestado"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcestado & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcparam) Then
            StrSql = StrSql & ",bprcparam"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcparam & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcprogreso) Then
            StrSql = StrSql & ",bprcprogreso"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcprogreso
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecfin) Then
            StrSql = StrSql & ",bprcfecfin"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecfin)
        End If
        If Not IsNull(rs_Batch_Proceso!bprchorafin) Then
            StrSql = StrSql & ",bprchorafin"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprchorafin & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprctiempo) Then
            StrSql = StrSql & ",bprctiempo"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprctiempo & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!Empnro) Then
            StrSql = StrSql & ",empnro"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!Empnro
        End If
        If Not IsNull(rs_Batch_Proceso!bprcpid) Then
            StrSql = StrSql & ",bprcPid"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcpid
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecInicioEj) Then
            StrSql = StrSql & ",bprcfecInicioEj"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecInicioEj)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecFinEj) Then
            StrSql = StrSql & ",bprcfecFinEj"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecFinEj)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcUrgente) Then
            StrSql = StrSql & ",bprcUrgente"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcUrgente
        End If
        If Not IsNull(rs_Batch_Proceso!bprcHoraInicioEj) Then
            StrSql = StrSql & ",bprcHoraInicioEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcHoraInicioEj & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcHoraFinEj) Then
            StrSql = StrSql & ",bprcHoraFinEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcHoraFinEj & "'"
        End If

        StrSql = StrSql & ") VALUES (" & StrSqlDatos & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'Reviso que haya copiado
        StrSql = "SELECT * FROM His_batch_proceso WHERE bpronro =" & NroProceso
        OpenRecordset StrSql, rs_His_Batch_Proceso
        
        If Not rs_His_Batch_Proceso.EOF Then
            ' Borro de Batch_proceso
            'StrSql = "DELETE FROM Batch_Proceso WHERE bpronro = " & NroProceso
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
        If rs_Batch_Proceso.State = adStateOpen Then rs_Batch_Proceso.Close
        If rs_His_Batch_Proceso.State = adStateOpen Then rs_His_Batch_Proceso.Close
    ' -----------------------------------------------------------------------------------


    If objConn.State = adStateOpen Then objConn.Close
    If objconnProgreso.State = adStateOpen Then objconnProgreso.Close
    
    Set rs_Batch_Proceso = Nothing
    Set rs_His_Batch_Proceso = Nothing

Final:
    Flog.writeline Espacios(Tabulador * 0) & "Fin de Actualización de categorias: " & " " & Now
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------------"
    Flog.Close
Exit Sub

CE:
    Flog.writeline Espacios(Tabulador * 0) & "Proceso abortado por Error"
    Flog.writeline Espacios(Tabulador * 1) & "Error:" & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    'MyRollbackTrans
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
End Sub




Public Function buscar_categoria(ByVal suma As Long) As Long
Dim aux As Long
Dim rs_buscategorias As New ADODB.Recordset

 StrSql = "SELECT estrnro FROM estructura WHERE  tenro=112 and (estrcodext ='" & suma
 StrSql = StrSql & "' OR estrcodext = '' ) ORDER BY estrcodext DESC"
 OpenRecordset StrSql, rs_buscategorias
    If Not rs_buscategorias.EOF Then
        buscar_categoria = rs_buscategorias!estrnro
    Else
       buscar_categoria = 0
    End If
  rs_buscategorias.Close
  

End Function

Public Function intersecta(ByVal FechaD1 As String, ByVal FechaH1 As String, ByVal FechaD2 As String, ByVal FechaH2 As String) As Boolean

Dim aux As Boolean
    aux = False
    If (InStr(1, FechaH1, "2199") > 0) And (InStr(1, FechaH2, "2199") > 0) Then
        aux = True
    Else
        If (InStr(1, FechaH1, "2199") > 0) And (FechaD1 <= FechaH2) Then
          aux = True
        Else
            If (InStr(1, FechaH2, "2199") > 0) And (FechaH1 >= FechaD2) Then
               aux = True
            Else
                If (FechaD1 >= FechaD2) And (FechaD1 <= FechaH2) Then
                   aux = True
                Else
                     If (FechaH1 >= FechaD1) And (FechaH1 <= FechaH2) Then
                      aux = True
                     End If
                End If
            End If
        End If
    End If
    intersecta = aux

End Function

Public Function sumar_Categoria(ByVal ternro As Long, ByVal Desde As String, ByVal Hasta As String) As Long

Dim Suma_cat As Long
Dim rs_subcategorias As New ADODB.Recordset
Dim hastaaux As String

Suma_cat = 0
'busco la estructura "Titulo" para el empleado
   Flog.writeline "busco la estructura Titulo para el empleado: " & ternro
    StrSql = "SELECT htetdesde,htethasta,estrcodext ,estructura.estrdabr FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
    StrSql = StrSql & "  Where (his_estructura.tenro = 105 Or his_estructura.tenro = 106 Or his_estructura.tenro = 107) And ternro = " & ternro
    StrSql = StrSql & " ORDER BY htetdesde,htethasta "
    OpenRecordset StrSql, rs_subcategorias
    Do While Not rs_subcategorias.EOF
        If EsNulo(rs_subcategorias!htethasta) Then
            hastaaux = "01/01/2199"
        Else
            hastaaux = rs_subcategorias!htethasta
        End If
        
        If intersecta(rs_subcategorias!htetdesde, hastaaux, Desde, Hasta) Then
           If rs_subcategorias!estrcodext <> "" And IsNumeric(rs_subcategorias!estrcodext) Then
                Suma_cat = Suma_cat + rs_subcategorias!estrcodext
           Else
                Flog.writeline "El codigo externo de la estructura es incorrecto"
           End If
        End If
        rs_subcategorias.MoveNext
        
    Loop
    rs_subcategorias.Close
    sumar_Categoria = Suma_cat
End Function


Public Sub CalculoCat(ternro As Long, FDesde As String, FHasta As String)
Dim Fechad As String  ' fecha desde de la subcategoría
Dim Fechah As String   ' fecha hasta de la subcategoría
Dim FechaAux As String
Dim suma As Long
Dim codCategoria As Integer
Dim htethastaaux As String
Dim UltimoHasta As String


Dim rs_categorias As New ADODB.Recordset
Dim rs_estr_categorias As New ADODB.Recordset

    
    
    On Error GoTo ME_Local
    
   ' Debug.Print ternro
    
   'busco la estructura "Categoria" para el empleado
   Flog.writeline "busco la estructura Categoria para el empleado: " & ternro
    StrSql = "SELECT * FROM his_estructura WHERE  tenro=112 and ternro =" & ternro
    StrSql = StrSql & " ORDER BY htetdesde,htethasta "
    OpenRecordset StrSql, rs_categorias
    Fechad = FDesde
    If FHasta = "" Then
         Fechah = "01/01/2199"
    Else
        Fechah = FHasta
    End If
    
    If rs_categorias.EOF Then '(1)
       
        suma = sumar_Categoria(ternro, FDesde, FHasta)
        If suma > 0 Then
                codCategoria = buscar_categoria(suma)
                If codCategoria <> 0 Then '(2)
                  
                        StrSql = "INSERT INTO his_estructura (tenro,ternro,estrnro,htetdesde,htethasta"
                        StrSql = StrSql & " ) VALUES ("
                        StrSql = StrSql & "112,"
                        StrSql = StrSql & ternro & ","
                        StrSql = StrSql & codCategoria & ","
                        StrSql = StrSql & ConvFecha(Fechad) & ","
                        If InStr(1, Fechah, "2199") > 0 Then
                           StrSql = StrSql & "null"
                        Else
                          StrSql = StrSql & ConvFecha(Fechah)
                          UltimoHasta = Fechah
                        End If
                        StrSql = StrSql & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline "8"
                    
                Else   '(2)
                    Flog.writeline "No existe ninguna estructura cuyo cod. externo sea: " & suma
                End If '(2)
        End If
    Else '(1)
         UltimoHasta = rs_categorias!htetdesde
        Do While (Not rs_categorias.EOF) And (CDate(Fechad) <= CDate(Fechah))
                If EsNulo(rs_categorias!htethasta) Then
                        htethastaaux = "01/01/2199"
                Else
                        htethastaaux = rs_categorias!htethasta
                End If
                If CDate(Fechad) <= CDate(htethastaaux) Then
                        If CDate(Fechad) = rs_categorias!htetdesde And CDate(Fechah) = htethastaaux Then 'intervalos iguales '(3)
                                  suma = sumar_Categoria(ternro, Fechad, Fechah)
                                  If suma > 0 Then
                                        codCategoria = buscar_categoria(suma)
                                        If codCategoria <> 0 Then  '(4)
                                              StrSql = " UPDATE his_estructura SET "
                                              StrSql = StrSql & " estrnro = " & codCategoria
                                              StrSql = StrSql & " WHERE hestrnro = " & rs_categorias!hestrnro
                                                 objConn.Execute StrSql, , adExecuteNoRecords
                                               Flog.writeline "u1"
                                              Fechad = DateAdd("d", 1, Fechah)
                                              UltimoHasta = Fechah
                                        Else '(4)
                                              Flog.writeline "No existe ninguna estructura cuyo cod. externo sea: " & suma
                                        End If '(4)
                                    Else
                                        StrSql = " DELETE FROM his_estructura "
                                         StrSql = StrSql & " WHERE hestrnro = " & rs_categorias!hestrnro
                                         objConn.Execute StrSql, , adExecuteNoRecords
                                         Fechad = DateAdd("d", 1, Fechah)
                                    End If
                        Else '(3)
                            If CDate(Fechad) = rs_categorias!htetdesde Then   'comienzo igual '(5)
                                If CDate(Fechah) < CDate(htethastaaux) Then   ' fecha hasta del nuevo intervalo mayor que el actual '(6)
                                    suma = sumar_Categoria(ternro, Fechad, Fechah)
                                    If suma > 0 Then
                                        codCategoria = buscar_categoria(suma)
                                        If codCategoria <> 0 Then '(7)
                                               StrSql = " UPDATE his_estructura SET "
                                               StrSql = StrSql & " estrnro = " & codCategoria
                                               If InStr(1, Fechah, "2199") > 0 Then
                                                    StrSql = StrSql & " ,htethasta = NULL"
                                               Else
                                                    StrSql = StrSql & " ,htethasta = " & ConvFecha(Fechah)
                                                    UltimoHasta = Fechah
                                               End If
                                               StrSql = StrSql & " WHERE hestrnro = " & rs_categorias!hestrnro
                                               Flog.writeline "u2"
                                                  objConn.Execute StrSql, , adExecuteNoRecords
                                               Fechad = DateAdd("d", 1, Fechah)
                                               Fechah = htethastaaux
                                               
                                         Else '(7)
                                               Flog.writeline "No existe ninguna estructura cuyo cod. externo sea: " & suma
                                         End If '(7)
                                    Else
                                        StrSql = " DELETE FROM his_estructura "
                                         StrSql = StrSql & " WHERE hestrnro = " & rs_categorias!hestrnro
                                         objConn.Execute StrSql, , adExecuteNoRecords
                                         Fechad = DateAdd("d", 1, Fechah)
                                         Fechah = htethastaaux
                                    End If
                                 Else   ' fecha hasta del nuevo intervalo menor que el actual '(6)
                                     suma = sumar_Categoria(ternro, Fechad, htethastaaux)
                                     If suma > 0 Then
                                           codCategoria = buscar_categoria(suma)
                                           If codCategoria <> 0 Then '(8)
                                                  StrSql = " UPDATE his_estructura SET "
                                                  StrSql = StrSql & " estrnro = " & codCategoria
                                                  StrSql = StrSql & " WHERE hestrnro = " & rs_categorias!hestrnro
                                                     objConn.Execute StrSql, , adExecuteNoRecords
                                                     Flog.writeline "u3"
                                                  Fechad = DateAdd("d", 1, htethastaaux)
                                                  UltimoHasta = htethastaaux
                                            Else '(8)
                                                  Flog.writeline "No existe ninguna estructura cuyo cod. externo sea: " & suma
                                            End If '(8)
                                     Else
                                        StrSql = " DELETE FROM his_estructura "
                                         StrSql = StrSql & " WHERE hestrnro = " & rs_categorias!hestrnro
                                         objConn.Execute StrSql, , adExecuteNoRecords
                                         Fechad = DateAdd("d", 1, htethastaaux)
                                    End If
                                 End If '(6)
                            Else   ' fecha desde de los intervalos distinta '(5)
                                   If rs_categorias!htetdesde > CDate(Fechad) Then  '(9)
                                         If htethastaaux > CDate(Fechah) Then  '(10)
                                            FechaAux = DateAdd("d", -1, rs_categorias!htetdesde)
                                            suma = sumar_Categoria(ternro, Fechad, FechaAux)
                                            codCategoria = buscar_categoria(suma)
                                            If codCategoria <> 0 Then  '(11)
                                                    StrSql = "INSERT INTO his_estructura (tenro,ternro,estrnro,htetdesde,htethasta"
                                                    StrSql = StrSql & " ) VALUES ("
                                                    StrSql = StrSql & "112,"
                                                    StrSql = StrSql & ternro & ","
                                                    StrSql = StrSql & codCategoria & ","
                                                    StrSql = StrSql & ConvFecha(Fechad) & ","
                                                    If InStr(1, FechaAux, "2199") > 0 Then
                                                        StrSql = StrSql & "null"
                                                    Else
                                                        StrSql = StrSql & ConvFecha(FechaAux)
                                                         UltimoHasta = FechaAux
                                                    End If
                                                    StrSql = StrSql & ")"
                                                objConn.Execute StrSql, , adExecuteNoRecords
                                                Flog.writeline "1"
                                               
                                            Else '(11)
                                                Flog.writeline "No existe ninguna estructura cuyo cod. externo sea: " & suma
                                            End If '(11)
                                            If CDate(Fechah) >= CDate(rs_categorias!htetdesde) Then
                                                suma = sumar_Categoria(ternro, rs_categorias!htetdesde, Fechah)
                                                If suma > 0 Then
                                                    codCategoria = buscar_categoria(suma)
                                                    If codCategoria <> 0 Then '(12)
                                                           StrSql = " UPDATE his_estructura SET "
                                                           StrSql = StrSql & " estrnro = " & codCategoria
                                                           If InStr(1, Fechah, "2199") > 0 Then
                                                              StrSql = StrSql & " ,htethasta = NULL"
                                                           Else
                                                                StrSql = StrSql & " ,htethasta = " & ConvFecha(Fechah)
                                                                UltimoHasta = Fechah
                                                           End If
                                                           StrSql = StrSql & " WHERE hestrnro = " & rs_categorias!hestrnro
                                                           Flog.writeline "u4"
                                                              objConn.Execute StrSql, , adExecuteNoRecords
                                                           Fechad = DateAdd("d", 1, Fechah)
                                                           Fechah = htethastaaux
                                                            
                                                     Else '(12)
                                                           Flog.writeline "No existe ninguna estructura cuyo cod. externo sea: " & suma
                                                     End If '(12)
                                                 Else
                                                      StrSql = " DELETE FROM his_estructura "
                                                      StrSql = StrSql & " WHERE hestrnro = " & rs_categorias!hestrnro
                                                      objConn.Execute StrSql, , adExecuteNoRecords
                                                     Fechad = DateAdd("d", 1, Fechah)
                                                     Fechah = htethastaaux
                                                    
                                                 End If
                                             Else
                                                Fechad = DateAdd("d", 1, Fechah)
                                             End If
                                         End If '(10)
                                         If CDate(htethastaaux) < CDate(Fechah) Then  '(10)
                                                FechaAux = DateAdd("d", -1, rs_categorias!htetdesde)
                                                suma = sumar_Categoria(ternro, Fechad, FechaAux)
                                                codCategoria = buscar_categoria(suma)
                                                If codCategoria <> 0 Then  '(11)
                                                        StrSql = "INSERT INTO his_estructura (tenro,ternro,estrnro,htetdesde,htethasta"
                                                        StrSql = StrSql & " ) VALUES ("
                                                        StrSql = StrSql & "112,"
                                                        StrSql = StrSql & ternro & ","
                                                        StrSql = StrSql & codCategoria & ","
                                                        StrSql = StrSql & ConvFecha(Fechad) & ","
                                                        If InStr(1, FechaAux, "2199") > 0 Then
                                                            StrSql = StrSql & "null"
                                                        Else
                                                            StrSql = StrSql & ConvFecha(FechaAux)
                                                            UltimoHasta = FechaAux
                                                        End If
                                                        StrSql = StrSql & ")"
                                                    objConn.Execute StrSql, , adExecuteNoRecords
                                                    Flog.writeline "2"
                                                     
                                                Else '(11)
                                                    Flog.writeline "No existe ninguna estructura cuyo cod. externo sea: " & suma
                                                End If '(11)
                                                suma = sumar_Categoria(ternro, rs_categorias!htetdesde, htethastaaux)
                                                If suma > 0 Then
                                                    codCategoria = buscar_categoria(suma)
                                                    If codCategoria <> 0 Then '(12)
                                                           StrSql = " UPDATE his_estructura SET "
                                                           StrSql = StrSql & " estrnro = " & codCategoria
                                                           StrSql = StrSql & " WHERE hestrnro = " & rs_categorias!hestrnro
                                                              objConn.Execute StrSql, , adExecuteNoRecords
                                                              Flog.writeline "u5"
                                                           Fechad = DateAdd("d", 1, htethastaaux)
                                                           UltimoHasta = htethastaaux
                                                     Else '(12)
                                                           Flog.writeline "No existe ninguna estructura cuyo cod. externo sea: " & suma
                                                     End If '(12)
                                                  Else
                                                        StrSql = " DELETE FROM his_estructura "
                                                         StrSql = StrSql & " WHERE hestrnro = " & rs_categorias!hestrnro
                                                         objConn.Execute StrSql, , adExecuteNoRecords
                                                        Fechad = DateAdd("d", 1, htethastaaux)
                                                  End If
                                         
                                         End If '(10)
                                         If CDate(htethastaaux) = CDate(Fechah) Then '(10)
                                                FechaAux = DateAdd("d", -1, rs_categorias!htetdesde)
                                                suma = sumar_Categoria(ternro, Fechad, FechaAux)
                                                codCategoria = buscar_categoria(suma)
                                                If codCategoria <> 0 Then  '(11)
                                                        StrSql = "INSERT INTO his_estructura (tenro,ternro,estrnro,htetdesde,htethasta"
                                                        StrSql = StrSql & " ) VALUES ("
                                                        StrSql = StrSql & "112,"
                                                        StrSql = StrSql & ternro & ","
                                                        StrSql = StrSql & codCategoria & ","
                                                        StrSql = StrSql & ConvFecha(Fechad) & ","
                                                        If InStr(1, FechaAux, "2199") > 0 Then
                                                            StrSql = StrSql & "null"
                                                        Else
                                                            StrSql = StrSql & ConvFecha(FechaAux)
                                                            UltimoHasta = FechaAux
                                                        End If
                                                        StrSql = StrSql & ")"
                                                    objConn.Execute StrSql, , adExecuteNoRecords
                                                    Flog.writeline "3"
                                                     
                                                Else '(11)
                                                    Flog.writeline "No existe ninguna estructura cuyo cod. externo sea: " & suma
                                                End If '(11)
                                                suma = sumar_Categoria(ternro, rs_categorias!htetdesde, htethastaaux)
                                                If suma > 0 Then
                                                    codCategoria = buscar_categoria(suma)
                                                    If codCategoria <> 0 Then '(12)
                                                           StrSql = " UPDATE his_estructura SET "
                                                           StrSql = StrSql & " estrnro = " & codCategoria
                                                           StrSql = StrSql & " WHERE hestrnro = " & rs_categorias!hestrnro
                                                              objConn.Execute StrSql, , adExecuteNoRecords
                                                              Flog.writeline "u6"
                                                           Fechad = DateAdd("d", 1, Fechah)
                                                           UltimoHasta = htethastaaux
                                                       
                                                     Else '(12)
                                                           Flog.writeline "No existe ninguna estructura cuyo cod. externo sea: " & suma
                                                     End If '(12)
                                                Else
                                                    StrSql = " DELETE FROM his_estructura "
                                                     StrSql = StrSql & " WHERE hestrnro = " & rs_categorias!hestrnro
                                                     objConn.Execute StrSql, , adExecuteNoRecords
                                                    Fechad = DateAdd("d", 1, Fechah)
                                                End If
                                         End If '(10)
                                   Else '(9)
                                          If htethastaaux > CDate(Fechah) Then  '(10)
                                            FechaAux = DateAdd("d", -1, Fechad)
                                            suma = sumar_Categoria(ternro, rs_categorias!htetdesde, FechaAux)
                                            codCategoria = buscar_categoria(suma)
                                            If codCategoria <> 0 Then  '(11)
                                                    StrSql = "INSERT INTO his_estructura (tenro,ternro,estrnro,htetdesde,htethasta"
                                                    StrSql = StrSql & " ) VALUES ("
                                                    StrSql = StrSql & "112,"
                                                    StrSql = StrSql & ternro & ","
                                                    StrSql = StrSql & codCategoria & ","
                                                    StrSql = StrSql & ConvFecha(rs_categorias!htetdesde) & ","
                                                    If InStr(1, FechaAux, "2199") > 0 Then
                                                        StrSql = StrSql & "null"
                                                    Else
                                                        StrSql = StrSql & ConvFecha(FechaAux)
                                                         UltimoHasta = FechaAux
                                                    End If
                                                    StrSql = StrSql & ")"
                                                objConn.Execute StrSql, , adExecuteNoRecords
                                                Flog.writeline "4"
                                                
                        
                                            Else '(11)
                                                Flog.writeline "No existe ninguna estructura cuyo cod. externo sea: " & suma
                                            End If '(11)
                                            suma = sumar_Categoria(ternro, Fechad, Fechah)
                                            If suma > 0 Then
                                                codCategoria = buscar_categoria(suma)
                                                If codCategoria <> 0 Then '(12)
                                                       StrSql = " UPDATE his_estructura SET "
                                                       StrSql = StrSql & " estrnro = " & codCategoria
                                                       StrSql = StrSql & " ,htetdesde = " & ConvFecha(Fechad)
                                                       If InStr(1, Fechah, "2199") > 0 Then
                                                          StrSql = StrSql & " ,htethasta = null"
                                                       Else
                                                            StrSql = StrSql & " ,htethasta = " & ConvFecha(Fechah)
                                                            UltimoHasta = Fechah
                                                       End If
                                                       StrSql = StrSql & " WHERE hestrnro = " & rs_categorias!hestrnro
                                                       Flog.writeline "u8"
                                                          objConn.Execute StrSql, , adExecuteNoRecords
                                                       Fechad = DateAdd("d", 1, Fechah)
                                                       Fechah = htethastaaux
                                                       
                                                 Else '(12)
                                                       Flog.writeline "No existe ninguna estructura cuyo cod. externo sea: " & suma
                                                 End If '(12)
                                            Else
                                                StrSql = " DELETE FROM his_estructura "
                                                 StrSql = StrSql & " WHERE hestrnro = " & rs_categorias!hestrnro
                                                 objConn.Execute StrSql, , adExecuteNoRecords
                                                 Fechad = DateAdd("d", 1, Fechah)
                                                 Fechah = htethastaaux
                                            End If
                                         End If '(10)
                                         If CDate(htethastaaux) < CDate(Fechah) Then  '(10)
                                                FechaAux = DateAdd("d", -1, Fechad)
                                                suma = sumar_Categoria(ternro, rs_categorias!htetdesde, FechaAux)
                                                codCategoria = buscar_categoria(suma)
                                                If codCategoria <> 0 Then  '(11)
                                                        StrSql = "INSERT INTO his_estructura (tenro,ternro,estrnro,htetdesde,htethasta"
                                                        StrSql = StrSql & " ) VALUES ("
                                                        StrSql = StrSql & "112,"
                                                        StrSql = StrSql & ternro & ","
                                                        StrSql = StrSql & codCategoria & ","
                                                        StrSql = StrSql & ConvFecha(rs_categorias!htetdesde) & ","
                                                        If InStr(1, FechaAux, "2199") > 0 Then
                                                            StrSql = StrSql & "null"
                                                        Else
                                                            StrSql = StrSql & ConvFecha(FechaAux)
                                                            UltimoHasta = FechaAux
                                                        End If
                                                        StrSql = StrSql & ")"
                                                    objConn.Execute StrSql, , adExecuteNoRecords
                                                    Flog.writeline "5"
                                                    
                                                Else '(11)
                                                    Flog.writeline "No existe ninguna estructura cuyo cod. externo sea: " & suma
                                                End If '(11)
                                                suma = sumar_Categoria(ternro, Fechad, htethastaaux)
                                                If suma > 0 Then
                                                    codCategoria = buscar_categoria(suma)
                                                    If codCategoria <> 0 Then '(12)
                                                           StrSql = " UPDATE his_estructura SET "
                                                           StrSql = StrSql & " estrnro = " & codCategoria
                                                           StrSql = StrSql & " ,htetdesde = " & ConvFecha(Fechad)
                                                           StrSql = StrSql & " WHERE hestrnro = " & rs_categorias!hestrnro
                                                              objConn.Execute StrSql, , adExecuteNoRecords
                                                              Flog.writeline "u9"
                                                           Fechad = DateAdd("d", 1, htethastaaux)
                                                          UltimoHasta = htethastaaux
                                                        
                                                     Else '(12)
                                                           Flog.writeline "No existe ninguna estructura cuyo cod. externo sea: " & suma
                                                     End If '(12)
                                                Else
                                                    StrSql = " DELETE FROM his_estructura "
                                                     StrSql = StrSql & " WHERE hestrnro = " & rs_categorias!hestrnro
                                                     objConn.Execute StrSql, , adExecuteNoRecords
                                                     Fechad = DateAdd("d", 1, htethastaaux)
                                                     
                                                End If
                                         
                                         End If '(10)
                                         If CDate(htethastaaux) = CDate(Fechah) Then '(10)
                                            UltimoHasta = DateAdd("d", 1, UltimoHasta)
                                             If CDate(UltimoHasta) < CDate(Fechad) Then
                                                FechaAux = rs_categorias!htetdesde
                                                Fechah = DateAdd("d", -1, Fechad) '**********modifiqué acá
                                             Else
                                               FechaAux = Fechad
                                             End If
                                                suma = sumar_Categoria(ternro, FechaAux, Fechah)
                                                codCategoria = buscar_categoria(suma)
                                                If codCategoria <> 0 Then  '(11)
                                                        StrSql = "INSERT INTO his_estructura (tenro,ternro,estrnro,htetdesde,htethasta"
                                                        StrSql = StrSql & " ) VALUES ("
                                                        StrSql = StrSql & "112,"
                                                        StrSql = StrSql & ternro & ","
                                                        StrSql = StrSql & codCategoria & ","
                                                        StrSql = StrSql & ConvFecha(FechaAux) & ","
                                                        If InStr(1, Fechah, "2199") > 0 Then
                                                            StrSql = StrSql & "null"
                                                        Else
                                                            StrSql = StrSql & ConvFecha(Fechah)
                                                        End If
                                                        StrSql = StrSql & ")"
                                                    objConn.Execute StrSql, , adExecuteNoRecords
                                                    Flog.writeline "6"
                                                    UltimoHasta = Fechah
                                                Else '(11)
                                                    Flog.writeline "No existe ninguna estructura cuyo cod. externo sea: " & suma
                                                End If '(11)
                                                 If CDate(UltimoHasta) < CDate(Fechad) Then
                                                        If CDate(Fechad) <= CDate(htethastaaux) Then
                                                            suma = sumar_Categoria(ternro, Fechad, htethastaaux)
                                                            If suma > 0 Then
                                                                codCategoria = buscar_categoria(suma)
                                                                If codCategoria <> 0 Then '(12)
                                                                       StrSql = " UPDATE his_estructura SET "
                                                                       StrSql = StrSql & " estrnro = " & codCategoria
                                                                         StrSql = StrSql & " ,htetdesde = " & ConvFecha(Fechad)
                                                                       StrSql = StrSql & " WHERE hestrnro = " & rs_categorias!hestrnro
                                                                          objConn.Execute StrSql, , adExecuteNoRecords
                                                                          Flog.writeline "u10"
                                                                       Fechad = DateAdd("d", 1, Fechah)
                                                                       UltimoHasta = htethastaaux
                                                                   
                                                                 Else '(12)
                                                                       Flog.writeline "No existe ninguna estructura cuyo cod. externo sea: " & suma
                                                                 End If '(12)
                                                             Else
                                                                StrSql = " DELETE FROM his_estructura "
                                                                 StrSql = StrSql & " WHERE hestrnro = " & rs_categorias!hestrnro
                                                                  objConn.Execute StrSql, , adExecuteNoRecords
                                                                Fechad = DateAdd("d", 1, Fechah)
                                                             End If
                                                         End If
                                                     Else
                                                     
                                                         If InStr(1, htethastaaux, "2199") > 0 Then
                                                            FechaAux = DateAdd("d", -1, Fechad)
                                                            If CDate(Fechad) <= CDate(htethastaaux) Then
                                                               suma = sumar_Categoria(ternro, rs_categorias!htetdesde, FechaAux)
                                                               If suma > 0 Then
                                                                   codCategoria = buscar_categoria(suma)
                                                                   If codCategoria <> 0 Then '(12)
                                                                          StrSql = " UPDATE his_estructura SET "
                                                                          StrSql = StrSql & " estrnro = " & codCategoria
                                                                            StrSql = StrSql & " ,htethasta = " & ConvFecha(FechaAux)
                                                                          StrSql = StrSql & " WHERE hestrnro = " & rs_categorias!hestrnro
                                                                             objConn.Execute StrSql, , adExecuteNoRecords
                                                                             Flog.writeline "u10"
                                                                          Fechad = DateAdd("d", 1, Fechah)
                                                                          UltimoHasta = htethastaaux
                                                                      
                                                                    Else '(12)
                                                                          Flog.writeline "No existe ninguna estructura cuyo cod. externo sea: " & suma
                                                                    End If '(12)
                                                                Else
                                                                   StrSql = " DELETE FROM his_estructura "
                                                                    StrSql = StrSql & " WHERE hestrnro = " & rs_categorias!hestrnro
                                                                    objConn.Execute StrSql, , adExecuteNoRecords
                                                                     Fechad = DateAdd("d", 1, Fechah)
                                                                End If
                                                            End If
                                                         
                                                         
                                                            
                                                         End If
                                                         Fechad = DateAdd("d", 1, Fechah)
                                                            
                                                     End If
                                         End If '(10)
                                   End If '(9)
                            End If '(5)
                        End If '(3)
                     '   objConn.Execute StrSql, , adExecuteNoRecords
             Else
                UltimoHasta = rs_categorias!htethasta
             End If
             rs_categorias.MoveNext
        
        Loop
       
        If CDate(Fechad) <= CDate(Fechah) Then '(13)
             suma = sumar_Categoria(ternro, Fechad, Fechah)
             codCategoria = buscar_categoria(suma)
            
             If codCategoria <> 0 Then  '(14)
               
                     StrSql = "INSERT INTO his_estructura (tenro,ternro,estrnro,htetdesde,htethasta"
                     StrSql = StrSql & " ) VALUES ("
                     StrSql = StrSql & "112,"
                     StrSql = StrSql & ternro & ","
                     StrSql = StrSql & codCategoria & ","
                     StrSql = StrSql & ConvFecha(Fechad) & ","
                    If InStr(1, Fechah, "2199") > 0 Then
                        StrSql = StrSql & "null"
                    Else
                        StrSql = StrSql & ConvFecha(Fechah)
                    End If
                     StrSql = StrSql & ")"
                 objConn.Execute StrSql, , adExecuteNoRecords
                 Flog.writeline "7"
                 UltimoHasta = Fechah
             Else '(14)
                 Flog.writeline "No existe ninguna estructura cuyo cod. externo sea: " & suma
             End If '(14)
        
        End If '(13)
    End If '(1)
   
    
   
ME_Local:


    Flog.writeline Espacios(Tabulador * 1) & " ---------------------------------------------------------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 2) & "SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & " ---------------------------------------------------------------------------------------------------"
    Flog.writeline
End Sub
