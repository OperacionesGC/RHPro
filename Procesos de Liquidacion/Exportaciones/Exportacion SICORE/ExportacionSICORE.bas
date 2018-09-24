Attribute VB_Name = "ExpSICORE"
' ---------------------------------------------------------------------------------------------
' Descripcion: Proyecto encargado de generar la exportacion de SICORE
' Autor      : GdeCos
' Fecha      : 29/05/2005
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------

Option Explicit

Global Const Version = "1.01"
Global Const FechaModificacion = "31/07/2009"
Global Const UltimaModificacion = "MB - Encriptacion de string connection"

Dim fs, f
'Global Flog

Dim NroProceso As Long

Global Path As String
Global HuboErrores As Boolean

Global sep As String
Global l_desde As Date
Global l_hasta As Date
Global l_incOperBen As Integer
Global l_tipo As Integer


Private Sub Main()

Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim StrSql2 As String
Dim objRs As New ADODB.Recordset
Dim objRs2 As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset
Dim fs2, fs3
Dim ArchExp, ArchExp2
Dim I As Integer
Dim PID As String
Dim parametros As String
Dim ArrParametros
Dim procesos_sep
Dim CantRegistros As Long
Dim Empresa As String
Dim Directorio As String
Dim Carpeta


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
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    On Error GoTo CE
    
    TiempoInicialProceso = GetTickCount
    'OpenConnection strconexion, objConn
    
    
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ExportacionSICORE" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    
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
    
    On Error GoTo CE
    
    Flog.writeline "Inicio Proceso de Exportación SICORE : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs2
    
    If Not objRs2.EOF Then
       
       'Obtengo los parametros del proceso
        ' desde@hasta@incOperBen@tipo@sep
        
        parametros = objRs2!bprcparam
        ArrParametros = Split(parametros, "@")
              
       'Obtengo la fecha inicial
        l_desde = CDate(ArrParametros(0))
        Flog.writeline "Fecha Desde: " & l_desde
       
       'Obtengo la fecha final
        l_hasta = CDate(ArrParametros(1))
        Flog.writeline "Fecha Hasta: " & l_hasta
        
        l_incOperBen = ArrParametros(2)
        If l_incOperBen = 0 Then
                Flog.writeline "No Incluye Operador"
        Else
                Flog.writeline "Incluye Operador"
        End If

       'Obtengo el tipo de reporte a exportar
        l_tipo = ArrParametros(3)
        If l_tipo = 0 Then
                Flog.writeline "Ambos tipos de reporte"
        Else
                If l_tipo = 1 Then
                        Flog.writeline "Exportacion de Reporte de SICORE"
                Else
                        Flog.writeline "Exportacion de Reporte de Retenidos"
                End If
        End If

       'Obtengo el separador decimal
        sep = ArrParametros(4)
        Flog.writeline "Separador seleccionado: " & sep
        
        
        'EMPIEZA EL PROCESO

        '------------------------------------------------------------------------------------------------------------------------
        'GENERO LA StrSql QUE BUSCA LOS DATOS
        '------------------------------------------------------------------------------------------------------------------------

           StrSql = " SELECT sicore.*, empresa.empnom, empresa.ternro, ter_doc.nrodoc "
           StrSql = StrSql & " FROM sicore "
           StrSql = StrSql & " INNER JOIN empresa ON empresa.estrnro = sicore.empresa "
           StrSql = StrSql & " INNER JOIN ter_doc ON ter_doc.ternro = empresa.ternro AND ter_doc.tidnro = 6 "
           StrSql = StrSql & " WHERE fec_desde = " & ConvFecha(l_desde) & " AND fec_hasta=" & ConvFecha(l_hasta)
           StrSql = StrSql & "   AND inc_oper_ben= " & l_incOperBen
           StrSql = StrSql & " ORDER BY ter_doc.nrodoc, empleg"

       OpenRecordset StrSql, objRs
        
       If objRs.EOF Then
          Flog.writeline "No se encontraron datos a Exportar."
               
       Else
                              
               'Directorio de exportacion
                StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
                OpenRecordset StrSql, rs
                If Not rs.EOF Then
                    Directorio = Trim(rs!sis_dirsalidas)
                End If
                
                StrSql = "SELECT * FROM modelo WHERE modnro = 259"
                OpenRecordset StrSql, rs_Modelo
                If Not rs_Modelo.EOF Then
                    If Not IsNull(rs_Modelo!modarchdefault) Then
                        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta desteino. El archivo será generado en el directorio default"
                    End If
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo. El archivo será generado en el directorio default"
                End If
                rs_Modelo.Close
                
                CantRegistros = CLng(objRs.RecordCount)
               
                Flog.writeline "Se comienza a procesar los datos"
               
                Select Case l_tipo
                        Case 1:
                               ' Genero los datos
                               Do Until objRs.EOF
                        
                                    Empresa = CStr(objRs!empnom)
                                                        
                                    Nombre_Arch = Directorio & "\" & CStr(objRs!nrodoc) & " - SICORE-" & Replace(CStr(l_desde), "/", "_") & "-" & Replace(CStr(l_hasta), "/", "_") & ".txt"
                                                    
                                    Flog.writeline "Se crea el archivo: " & Nombre_Arch
                                    ' Se Crea el archivo con nombre de la empresa y estrnro
                                    Set fs2 = CreateObject("Scripting.FileSystemObject")
                                    'Activo el manejador de errores
                                    On Error Resume Next
                                    If Err.Number <> 0 Then
                                        Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
                                        Set Carpeta = fs2.CreateFolder(Directorio)
                                    End If
                                    'desactivo el manejador de errores
                                    Set ArchExp = fs2.CreateTextFile(Nombre_Arch, True)
                                    On Error GoTo CE
                    
                                    I = 1
                    
                                    Do While (CStr(objRs!empnom) = Empresa)
                                    
                                        Flog.writeline "Se escribe el registro Nro.: " & I & " de la Empresa: " & Empresa
                                        
                                        ArchExp.Write "0"
                                        Call imprimirTexto(objRs!signo, ArchExp, 1, True)
                                        Call imprimirTexto(objRs!fec_ret, ArchExp, 10, True)
                                        Call imprimirNumero(objRs!cod_liq, ArchExp, 16)
                                        Call imprimirNumero(objRs!neto_ent, ArchExp, 13)
                                        ArchExp.Write sep
                                        Call imprimirNumero(objRs!neto_dec, ArchExp, 2)
                                        Call imprimirTexto(objRs!cod_impuesto, ArchExp, 3, True)
                                        Call imprimirTexto(objRs!cod_regimen, ArchExp, 3, True)
                                        ArchExp.Write "1"
                                        Call imprimirNumero(objRs!impo_ent, ArchExp, 11)
                                        ArchExp.Write sep
                                        Call imprimirNumero(objRs!impo_dec, ArchExp, 2)
                                        Call imprimirTexto(objRs!fec_ret, ArchExp, 10, True)
                                        Call imprimirTexto(objRs!cod_condicion, ArchExp, 2, True)
                                        Call imprimirNumero(objRs!gan_ent, ArchExp, 11)
                                        ArchExp.Write sep
                                        Call imprimirNumero(objRs!gan_dec, ArchExp, 2)
                                        Call imprimirTexto(cambiarFormato(objRs!porcexl, sep), ArchExp, 6, True)
                                        Call imprimirTexto(objRs!fec_bol, ArchExp, 10, True)
                                        Call imprimirTexto(objRs!tipo_doc, ArchExp, 2, True)
                                        Call imprimirTexto(objRs!Cuil, ArchExp, 11, True)
                                        ArchExp.Write "         "
                                        Call imprimirNumero(objRs!cod_liq, ArchExp, 14)
                                        If CInt(l_incOperBen) = -1 Then
                                           Call imprimirTexto(objRs!deno_orden, ArchExp, 30, True)
                                           Call imprimirTexto(objRs!acrecent, ArchExp, 1, True)
                                           Call imprimirTexto(objRs!cuit_pais, ArchExp, 11, True)
                                           Call imprimirTexto(objRs!cuit_orden, ArchExp, 11, True)
                                        End If
                                        
                                                                      
                                        ArchExp.writeline ""
                                                      
                                        TiempoAcumulado = GetTickCount
                                          
                                        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix((I / CantRegistros) * 100) & _
                                                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                                                 " WHERE bpronro = " & NroProceso
                                        objConn.Execute StrSql, , adExecuteNoRecords
                                         
                                        I = I + 1
                                                            
                                        objRs.MoveNext
                                
                                        If objRs.EOF Then
                                             Exit Do
                                        End If
                                                             
                                    Loop
                                                         
                                    ArchExp.Close
                                                         
                               Loop
                
                        Case 2:
                               ' Genero los datos
                               Do Until objRs.EOF
                        
                                    Empresa = CStr(objRs!empnom)
                                                        
                                    Nombre_Arch = Directorio & "\" & CStr(objRs!nrodoc) & " - SICORE-" & Replace(CStr(l_desde), "/", "_") & "-" & Replace(CStr(l_hasta), "/", "_") & ".txt"
                                                    
                                    Flog.writeline "Se crea el archivo: " & Nombre_Arch
                                    ' Se Crea el archivo con nombre de la empresa y estrnro
                                    Set fs2 = CreateObject("Scripting.FileSystemObject")
                                    'Activo el manejador de errores
                                    On Error Resume Next
                                    If Err.Number <> 0 Then
                                        Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
                                        Set Carpeta = fs2.CreateFolder(Directorio)
                                    End If
                                    'desactivo el manejador de errores
                                    Set ArchExp = fs2.CreateTextFile(Nombre_Arch, True)
                                    On Error GoTo CE
                    
                                    I = 1
                    
                                    Do While (CStr(objRs!empnom) = Empresa)
                                    
                                        Flog.writeline "Se escribe el registro Nro.: " & I & " de la Empresa: " & Empresa
                                        
                                        Call imprimirNumero(objRs!Cuil, ArchExp, 11)
                                        Call imprimirTexto(objRs!apenom, ArchExp, 20, True)
                                        Call imprimirTexto(objRs!dom_fiscal, ArchExp, 20, True)
                                        Call imprimirTexto(objRs!dom_localidad, ArchExp, 20, True)
                                        Call imprimirTexto(objRs!cod_provincia, ArchExp, 2, True)
                                        Call imprimirTexto(objRs!dom_cp, ArchExp, 8, False)
                                        Call imprimirTexto(objRs!tipo_doc, ArchExp, 2, True)
                                                                      
                                        ArchExp.writeline ""
                                                      
                                        TiempoAcumulado = GetTickCount
                                          
                                        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix((I / CantRegistros) * 100) & _
                                                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                                                 " WHERE bpronro = " & NroProceso
                                        objConn.Execute StrSql, , adExecuteNoRecords
                                         
                                        I = I + 1
                                                            
                                        objRs.MoveNext
                                
                                        If objRs.EOF Then
                                             Exit Do
                                        End If
                                                             
                                    Loop
                                                         
                                    ArchExp.Close
                                                         
                               Loop

                        Case 0:
                               ' Genero los datos
                               Do Until objRs.EOF
                        
                                    'Empresa
                                    Empresa = CStr(objRs!empnom)
                                                        
                                    'Genero el archivo de SICORE
                                    Nombre_Arch = Directorio & "\" & CStr(objRs!nrodoc) & " - SICORE-" & Replace(CStr(l_desde), "/", "_") & "-" & Replace(CStr(l_hasta), "/", "_") & ".txt"
                                                    
                                    Flog.writeline "Se crea el archivo: " & Nombre_Arch
                                    ' Se Crea el archivo con nombre de la empresa y estrnro
                                    Set fs2 = CreateObject("Scripting.FileSystemObject")
                                    'Activo el manejador de errores
                                    On Error Resume Next
                                    If Err.Number <> 0 Then
                                        Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
                                        Set Carpeta = fs2.CreateFolder(Directorio)
                                    End If
                                    'desactivo el manejador de errores
                                    Set ArchExp = fs2.CreateTextFile(Nombre_Arch, True)
                                    On Error GoTo CE
                    
                                    'Fin generacion archivo SICORE
                                    
                                    'Genero el archivo de Retenidos
                                    Nombre_Arch = Directorio & "\" & CStr(objRs!nrodoc) & "- Retenidos-" & Replace(CStr(l_desde), "/", "_") & "-" & Replace(CStr(l_hasta), "/", "_") & ".txt"
                                                    
                                    Flog.writeline "Se crea el archivo: " & Nombre_Arch
                                    ' Se Crea el archivo con nombre de la empresa y estrnro
                                    Set fs3 = CreateObject("Scripting.FileSystemObject")
                                    'Activo el manejador de errores
                                    On Error Resume Next
                                    If Err.Number <> 0 Then
                                        Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
                                        Set Carpeta = fs3.CreateFolder(Directorio)
                                    End If
                                    'desactivo el manejador de errores
                                    Set ArchExp2 = fs3.CreateTextFile(Nombre_Arch, True)
                                    On Error GoTo CE
                                    
                                    'Fin Generacion Archivo Retenidos
                                    
                                    I = 1
                                    'Comienzo a generar los datos de los dos archivos
                                    Do While (CStr(objRs!empnom) = Empresa)
                                    
                                        Flog.writeline "Se escribe el registro Nro.: " & I & " de la Empresa: " & Empresa
                                        Flog.writeline "En el archivo de SICORE"
                                        
                                        ArchExp.Write "0"
                                        Call imprimirTexto(objRs!signo, ArchExp, 1, True)
                                        Call imprimirTexto(objRs!fec_ret, ArchExp, 10, True)
                                        Call imprimirNumero(objRs!cod_liq, ArchExp, 16)
                                        Call imprimirNumero(objRs!neto_ent, ArchExp, 13)
                                        ArchExp.Write sep
                                        Call imprimirNumero(objRs!neto_dec, ArchExp, 2)
                                        Call imprimirTexto(objRs!cod_impuesto, ArchExp, 3, True)
                                        Call imprimirTexto(objRs!cod_regimen, ArchExp, 3, True)
                                        ArchExp.Write "1"
                                        Call imprimirNumero(objRs!impo_ent, ArchExp, 11)
                                        ArchExp.Write sep
                                        Call imprimirNumero(objRs!impo_dec, ArchExp, 2)
                                        Call imprimirTexto(objRs!fec_ret, ArchExp, 10, True)
                                        Call imprimirTexto(objRs!cod_condicion, ArchExp, 2, True)
                                        Call imprimirNumero(objRs!gan_ent, ArchExp, 11)
                                        ArchExp.Write sep
                                        Call imprimirNumero(objRs!gan_dec, ArchExp, 2)
                                        Call imprimirTexto(cambiarFormato(objRs!porcexl, sep), ArchExp, 6, True)
                                        Call imprimirTexto(objRs!fec_bol, ArchExp, 10, True)
                                        Call imprimirTexto(objRs!tipo_doc, ArchExp, 2, True)
                                        Call imprimirTexto(objRs!Cuil, ArchExp, 11, True)
                                        ArchExp.Write "         "
                                        Call imprimirNumero(objRs!cod_liq, ArchExp, 14)
                                        If CInt(l_incOperBen) = -1 Then
                                           Call imprimirTexto(objRs!deno_orden, ArchExp, 30, True)
                                           Call imprimirTexto(objRs!acrecent, ArchExp, 1, True)
                                           Call imprimirTexto(objRs!cuit_pais, ArchExp, 11, True)
                                           Call imprimirTexto(objRs!cuit_orden, ArchExp, 11, True)
                                        End If
                                                                      
                                        ArchExp.writeline ""
                                                      
                                        
                                        Flog.writeline "Se escribe el registro Nro.: " & I & " de la Empresa: " & Empresa
                                        Flog.writeline "En el archivo de Retenidos"
                                        
                                        Call imprimirNumero(objRs!Cuil, ArchExp2, 11)
                                        Call imprimirTexto(objRs!apenom, ArchExp2, 20, True)
                                        Call imprimirTexto(objRs!dom_fiscal, ArchExp2, 20, True)
                                        Call imprimirTexto(objRs!dom_localidad, ArchExp2, 20, True)
                                        Call imprimirTexto(objRs!cod_provincia, ArchExp2, 2, True)
                                        Call imprimirTexto(objRs!dom_cp, ArchExp2, 8, False)
                                        Call imprimirTexto(objRs!tipo_doc, ArchExp2, 2, True)
                                                                      
                                        ArchExp2.writeline ""
                                                      
                                        TiempoAcumulado = GetTickCount
                                          
                                        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix((I / CantRegistros) * 100) & _
                                                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                                                 " WHERE bpronro = " & NroProceso
                                        objConn.Execute StrSql, , adExecuteNoRecords
                                         
                                        I = I + 1
                                                            
                                        objRs.MoveNext
                                
                                        If objRs.EOF Then
                                             Exit Do
                                        End If
                                                             
                                    Loop
                                                         
                                    ArchExp.Close
                                    ArchExp2.Close
                                                         
                               Loop
                
                End Select
                        
          Flog.writeline "Se Terminaron de Procesar los datos"
       
       End If
    
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
    If objRs.State = adStateOpen Then objRs.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now

End Sub

Function cambiarFormato(cadena, separador)

 Dim l_salida
 
 l_salida = cadena

 If (InStr(cadena, ".")) Then
    If separador = "," Then
       l_salida = Replace(cadena, ".", ",")
    End If
 Else
     If (InStr(cadena, ",")) Then
        If separador = "." Then
           l_salida = Replace(cadena, ",", ".")
        End If
     End If
 End If

 cambiarFormato = l_salida

End Function 'cambiarFormato(cadena,separador)

Sub imprimirTexto(Texto, archivo, Longitud, derecha)
'Rutina genérica para imprimir un TEXTO, de una LONGITUD determinada.
'Los sobrantes se rellenan con CARACTER

Dim cadena
Dim u
Dim longTexto
    
    If IsNull(Texto) Then
        longTexto = 0
        cadena = " "
    Else
        longTexto = Len(Texto)
        cadena = CStr(Texto)
    End If
    
    u = CInt(Longitud) - longTexto
    If u < 0 Then
        archivo.Write cadena
    Else
        If derecha Then
            archivo.Write cadena & String(u, " ")
        Else
            archivo.Write String(u, " ") & cadena
        End If
    End If

End Sub

Sub imprimirNumero(Texto, archivo, Longitud)
'Rutina genérica para imprimir un TEXTO, de una LONGITUD determinada.
'Los sobrantes se rellenan con CARACTER

Dim cadena
Dim u
Dim longTexto
    
    If IsNull(Texto) Then
        longTexto = 0
        cadena = "0"
    Else
        longTexto = Len(Texto)
        cadena = CStr(Texto)
    End If
    
    u = CInt(Longitud) - longTexto
    If u < 0 Then
        archivo.Write cadena
    Else
        archivo.Write String(u, "0") & cadena
    End If

End Sub

