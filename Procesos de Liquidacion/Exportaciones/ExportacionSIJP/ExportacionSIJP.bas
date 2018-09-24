Attribute VB_Name = "ExpSIJP"
' ---------------------------------------------------------------------------------------------
' Descripcion: Proyecto encargado de generar la exportacion de SIJP
' Autor      : GdeCos
' Fecha      : 26/05/2005
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------

Option Explicit

Global Const Version = "1.01"
Global Const FechaModificacion = "31/07/2009"
Global Const UltimaModificacion = " " 'MB - Encriptacion de string connection

Dim fs, f
'Global Flog

Dim NroProceso As Long

Global Path As String
Global HuboErrores As Boolean

Global pliqnro As Long
Global lista_pronro As String
Global sep As String



Private Sub Main()

Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim StrSql2 As String
Dim objRs As New ADODB.Recordset
Dim objRs2 As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset
Dim fs2
Dim ArchExp
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
    
    Nombre_Arch = PathFLog & "ExportacionSIJP" & "-" & NroProceso & ".log"
    
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
    
    On Error GoTo CE
    
    Flog.writeline "Inicio Proceso de Exportación SIJP : " & Now
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
       parametros = objRs2!bprcparam
       ArrParametros = Split(parametros, "@")
              
       'Obtengo el Periodo
        pliqnro = ArrParametros(0)
        Flog.writeline "Periodo: " & pliqnro
       
       'Obtengo la lista de procesos
        lista_pronro = ArrParametros(1)
        Flog.writeline "Procesos: " & lista_pronro
        
       'Obtengo los datos del separador
        sep = ArrParametros(2)
        Flog.writeline "Separador seleccionado: " & sep
       
        
        '------------------------------------------------------------------------------------------------------------------------
        'GENERO LA StrSql QUE BUSCA LOS DATOS
        '------------------------------------------------------------------------------------------------------------------------

        StrSql = " SELECT repsijp.*, empresa.empnom, periodo.pliqdesc, empresa.ternro, ter_doc.nrodoc "
        StrSql = StrSql & " FROM repsijp "
        StrSql = StrSql & " INNER JOIN empresa ON empresa.estrnro = repsijp.empresa "
        StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = repsijp.pliqnro"
        StrSql = StrSql & " INNER JOIN ter_doc ON ter_doc.ternro = empresa.ternro AND ter_doc.tidnro = 6 "
        StrSql = StrSql & " WHERE repsijp.pliqnro = " & pliqnro & " AND repsijp.lista_pronro = '" & lista_pronro & "' "
        StrSql = StrSql & " ORDER BY ter_doc.nrodoc "
        
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
                
                StrSql = "SELECT * FROM modelo WHERE modnro = 258"
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
               
               ' Genero los datos
               Do Until objRs.EOF
        
                    Empresa = CStr(objRs!empnom)
                                        
                    Nombre_Arch = Directorio & "\" & CStr(objRs!nrodoc) & "_SIJP.txt"
                                    
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
                        
                        Call imprimirTexto(objRs!Cuil, ArchExp, 11, True)
                        Call imprimirTexto(objRs!apenom, ArchExp, 30, True)
                        Call imprimirTexto(objRs!Conyuges, ArchExp, 1, True)  'Conyugue
                        Call imprimirTexto(objRs!cant_hijos, ArchExp, 2, False) 'Cantidad de Hjos
                        Call imprimirTexto(objRs!cod_sitr, ArchExp, 2, False) 'Codigo de Situacion
                        
                        Call imprimirTexto(objRs!cod_cond, ArchExp, 2, False) 'Codigo de Condicion
                        Call imprimirTexto(objRs!actividad, ArchExp, 2, False) 'Codigo de Actividad
                        Call imprimirTexto(objRs!zona, ArchExp, 2, False)     'Codigo de Zona
                        Call imprimirTexto(objRs!PORC_ADI, ArchExp, 5, False) 'Porcentaje de Aporte adicional SS
                        Call imprimirTexto(objRs!cod_cont, ArchExp, 3, False) 'Codigo de Contrato
                        
                        Call imprimirTexto(objRs!cod_obra_social, ArchExp, 6, False) 'Codigo de Obra Social
                        Call imprimirTexto(objRs!adherentes, ArchExp, 2, False) 'Cantidad de Adherentes
                        Call imprimirTexto(cambiarFormato(objRs!rem_total, sep), ArchExp, 9, False) 'Remuneracion Total
                        Call imprimirTexto(cambiarFormato(objRs!imp_ss, sep), ArchExp, 9, False)  'Remuneracion Imponible
                        Call imprimirTexto(cambiarFormato(objRs!asig_fliar, sep), ArchExp, 9, False) 'Asignaciones Familiares Pagadas
                        
                        Call imprimirTexto(cambiarFormato(objRs!aporte_voluntario, sep), ArchExp, 9, False) 'Aporte Voluntario
                        Call imprimirTexto(cambiarFormato(objRs!adi_os, sep), ArchExp, 9, False)  'Adicional OS
                        Call imprimirTexto(cambiarFormato(objRs!exc_ss, sep), ArchExp, 9, False)  'Excedentes Aportes SS
                        Call imprimirTexto(cambiarFormato(objRs!exc_os, sep), ArchExp, 9, False)  'Excedentes Aportes OS
                        Call imprimirTexto(objRs!Localidad, ArchExp, 50, False) 'Provincia Localidad
                        
                        Call imprimirTexto(cambiarFormato(objRs!imp_ss_con, sep), ArchExp, 9, False) 'Remuneracion Imponible SS Contribuciones
                        
                        'SIJP Version 14
                        Call imprimirTexto(cambiarFormato(objRs!imprem3, sep), ArchExp, 9, False) 'Remineracion Imponible3 Tope 75 Mopres
                        Call imprimirTexto(cambiarFormato(objRs!imprem4, sep), ArchExp, 9, False) 'Remineracion Imponible4 Tope 60 Mopres
                        Call imprimirTexto(objRs!codsiniestro, ArchExp, 2, False) 'Codigo Siniestrado
                        Call imprimirTexto(objRs!correspred, ArchExp, 1, False) 'Corresp. Reduccion
                        Call imprimirTexto(cambiarFormato(objRs!caprecomlrt, sep), ArchExp, 9, False) 'Capital Recomp. de LRT
                        
                        'SIJP Version 17
                        Call imprimirTexto(objRs!tipempnro, ArchExp, 1, False) 'Codigo de DGI del tipo de empleador
                        Call imprimirTexto(cambiarFormato(objRs!apo_adi_os, sep), ArchExp, 9, False) 'Aportes adicionales OS
                        'Call imprimirTexto(objRs!Con_FNE, ArchExp, 1, False)  'Regimen Reparto o AFJP
                        Call imprimirTexto(objRs!secrep, ArchExp, 1, False)  'Regimen Reparto o AFJP
                        
                        'SIJP Version 20
                        Call imprimirTexto(objRs!cod_sitr1, ArchExp, 2, False) 'Codigo de Situacion de Revista 1
                        Call imprimirTexto(objRs!diainisr1, ArchExp, 2, False) 'Dia inicio SR 1
                        Call imprimirTexto(objRs!cod_sitr2, ArchExp, 2, False) 'Codigo de Situacion de Revista 2
                        Call imprimirTexto(objRs!diainisr2, ArchExp, 2, False) 'Dia inicio SR 2
                        Call imprimirTexto(objRs!cod_sitr3, ArchExp, 2, False) 'Codigo de Situacion de Revista 3
                        Call imprimirTexto(objRs!diainisr3, ArchExp, 2, False) 'Dia inicio SR 3
                        
                        Call imprimirTexto(cambiarFormato(objRs!sue_adic, sep), ArchExp, 9, False) 'Sueldo + Adicionales
                        Call imprimirTexto(cambiarFormato(objRs!sac, sep), ArchExp, 9, False)     'SAC
                        Call imprimirTexto(cambiarFormato(objRs!hrsextras, sep), ArchExp, 9, False) 'Horas Extras
                        Call imprimirTexto(cambiarFormato(objRs!zonadesf, sep), ArchExp, 9, False) 'Zona Desfavorable
                        Call imprimirTexto(cambiarFormato(objRs!lar, sep), ArchExp, 9, False)   'Vacaciones
                        
                        Call imprimirTexto(cambiarFormato(objRs!diastrab, sep), ArchExp, 9, False) 'Dias Trabajados
                        Call imprimirTexto(cambiarFormato(objRs!imprem5, sep), ArchExp, 9, False) 'Remuneracion Imponible 5
                        Call imprimirTexto(objRs!xconv, ArchExp, 1, False)    'Trabajador Convenciodado
                                                      
                        'SIJP Version 24
                        Call imprimirTexto(cambiarFormato("0.00", sep), ArchExp, 9, False) 'Remuneracion Imponible 6
                        Call imprimirTexto(cambiarFormato("0", sep), ArchExp, 1, False) 'Tipo de operacion
                        
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
Dim txt
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

