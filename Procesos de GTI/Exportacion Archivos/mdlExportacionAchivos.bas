Attribute VB_Name = "mdlExportacionAchivos"
Option Explicit
'Modulo para procedimientos de exportacion

Function modelo_2006(ByVal modnro As Long, ByVal archSalida As String)

Dim rs_datos  As New ADODB.Recordset
Dim Progreso As Double
Dim porc As Double
Dim cantEmpleados As Long
Dim fs
Dim strlinea As String
Dim archExp

    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Comienza la exportacion del modelo: " & modnro

    On Error GoTo ce
        
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Right(directorio, 1) = "\" Then
        archSalida = directorio & archSalida
    Else
        archSalida = directorio & "\" & archSalida
    End If
        
    StrSql = " SELECT empleado.empleg, empleado.terape, empleado.ternom, ter_doc.nrodoc,gti_histarjeta.hstjnrotar " & _
             " , gti_histarjeta.hstjfecdes, fases.altfec  " & _
             " FROM Empleado " & _
             " INNER JOIN tercero on tercero.ternro = empleado.ternro " & _
             " INNER JOIN tipodocu_pais on tipodocu_pais.paisnro = tercero.docpais AND tipodocu_pais.tidcod <= 4 " & _
             " INNER JOIN ter_doc on ter_doc.tidnro = tipodocu_pais.tidnro AND ter_doc.ternro = empleado.ternro " & _
             " INNER JOIN gti_histarjeta on gti_histarjeta.ternro = empleado.ternro " & _
             " INNER JOIN fases on fases.empleado = empleado.ternro " & _
             " WHERE empleado.empest = -1 AND (gti_histarjeta.hstjfecdes <= " & ConvFecha(Date) & " AND (gti_histarjeta.hstjfechas >= " & ConvFecha(Date) & " OR gti_histarjeta.hstjfechas is null)) " & _
             " AND (fases.altfec <= " & ConvFecha(Date) & " AND (fases.bajfec >= " & ConvFecha(Date) & " OR fases.bajfec is null)) " & _
             " ORDER BY empleado.empleg ASC "
    OpenRecordset StrSql, rs_datos
    
    cantEmpleados = rs_datos.RecordCount
    
    If cantEmpleados = 0 Then
        Progreso = 100
        cantEmpleados = 1
        Flog.writeline Espacios(Tabulador * 0) & "No se encontraron empleados."
    Else
        Flog.writeline Espacios(Tabulador * 0) & "Cantidad de registros a generar: " & cantEmpleados
        If fs.FileExists(archSalida) Then
            Flog.writeline Espacios(Tabulador * 0) & "Existe un archivo en el directorio se borrara"
            fs.deletefile archSalida, True
            Flog.writeline Espacios(Tabulador * 0) & "Se borro el archivo"
        End If
        Set archExp = fs.CreateTextFile(archSalida, True)
    End If
    porc = 100 / CLng(cantEmpleados)
    
    If usaencabezado Then
        strlinea = "ID" & separador & "NAME" & separador & "LASTNAME" & separador & "NAMEEMPLOYEE" & separador & "REGISTERSYSTEMDATE" & separador & "ACTIVEDAYS" & separador & "EMPLOYEECODE"
        archExp.writeline strlinea
    End If
    
    Do While Not rs_datos.EOF
    
        Progreso = Progreso + porc
        
        strlinea = ""
        strlinea = strlinea & rs_datos!empleg                                                  'ID
        strlinea = strlinea & separador & """" & rs_datos!nrodoc & """"                        'Name (DNI)
        strlinea = strlinea & separador & """" & rs_datos!terape & """"                        'LastName
        strlinea = strlinea & separador & """" & rs_datos!ternom & """"                        'NAMEEMPLOYEE
        strlinea = strlinea & separador & FormatoFecha(rs_datos!hstjfecdes, "AAAA-MM-DD")        'RegistersystemDate
        strlinea = strlinea & separador & FormatoFecha(rs_datos!altfec, "AAAA-MM-DD")            'Activeday
        strlinea = strlinea & separador & """" & rs_datos!empleg & """"                        'EMPLOYEECODE
        
        archExp.writeline strlinea
        
        Flog.writeline Espacios(Tabulador * 0) & "Se genero el registro para el legajo: " & rs_datos!empleg
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
        

        rs_datos.MoveNext
    Loop
        
    GoTo datosOK
ce:
    'strLinea = ""
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & "SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Incompleto = True
    Exit Function
datosOK:
    
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Fin exportacion del modelo: " & modnro

    Progreso = 100
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Set rs_datos = Nothing
End Function

