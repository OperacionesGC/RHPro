Attribute VB_Name = "mdlExportacionBDO"
Option Explicit

Function expEmpleado(ByVal ternro As Long, ByVal separador As String)
 Dim rsDatosEmp  As New ADODB.Recordset
 Dim rsAux  As New ADODB.Recordset
 Dim strLinea As String

    On Error GoTo CE
    Flog.writeline Espacios(Tabulador * 1) & "Generando la exportacion para el ternro: " & ternro & "."
    StrSql = "SELECT empleg,tercero.ternom,tercero.ternom2,tercero.terape,tercero.terape2,tercero.terfecnac,pais.paisdesc " & _
            ",tercero.terfecing,estcivil.estcivdesabr,tercero.tersex,empleado.empfecalta,empleado.empemail, dni.nrodoc dni, cuil.nrodoc cuil" & _
            ",empleado.empfbajaprev,empleado.empest,empleado.empremu,empleado.empreporta " & _
            " FROM empleado " & _
            " INNER JOIN tercero ON tercero.ternro = empleado.ternro  " & _
            " LEFT JOIN pais ON pais.paisnro = tercero.paisnro " & _
            " INNER JOIN estcivil ON estcivil.estcivnro = tercero.estcivnro " & _
            " LEFT JOIN ter_doc dni ON dni.ternro = empleado.ternro AND dni.tidnro = 1 " & _
            " LEFT JOIN ter_doc cuil ON cuil.ternro = empleado.ternro AND cuil.tidnro = 10 " & _
            " WHERE empleado.ternro = " & ternro
    OpenRecordset StrSql, rsDatosEmp
    If Not rsDatosEmp.EOF Then
        'legajo y empresa
        strLinea = rsDatosEmp!empleg                                'pos 1 legajo
        strLinea = strLinea & separador & obtenerEstructura(ternro, 10, Date)     'pos 2 empresa
        strLinea = strLinea & separador                             'pos 3 vacio
        'primer y segundo nombre
        strLinea = strLinea & separador & rsDatosEmp!ternom         'pos 4 nombre
        If Not IsNull(rsDatosEmp!ternom2) Then
            strLinea = strLinea & " " & rsDatosEmp!ternom2
        End If
        
        strLinea = strLinea & separador                             'pos 5 vacio
        strLinea = strLinea & separador                             'pos 6 vacio
        strLinea = strLinea & separador                             'pos 7 vacio
        
        'apellido
        'strLinea = strLinea & separador & rsDatosEmp!terape    'VER si se informa
        'sexo
        If (CLng(rsDatosEmp!tersex) = -1) Then
            strLinea = strLinea & separador & "M"                   'pos 8 sexo
        Else
            strLinea = strLinea & separador & "F"
        End If
        'fecha de nacimiento
        strLinea = strLinea & separador & rsDatosEmp!terfecnac      'pos 9 fecha de nacimiento
        'estado civil
        strLinea = strLinea & separador & rsDatosEmp!estcivdesabr   'pos 10 estado civil
        'pais de nacimiento
        strLinea = strLinea & separador & rsDatosEmp!paisDesc       'pos 11 pais de nacimiento
        'fecha de ingreso
        If EsNulo(rsDatosEmp!terfecing) Then
            strLinea = strLinea & separador
        Else
            strLinea = strLinea & separador & rsDatosEmp!terfecing  'pos 12 fecha de alta
            'strLinea = strLinea & separador & rsDatosEmp!empfecalta
        End If
        'fecha de fase mas antigua
        strLinea = strLinea & separador & calcularFase(ternro, separador)       'pos 13 fecha de alta
        
        'strLinea = strLinea & separador & obtenerEstructura(ternro, 1, Date)   'pos 14 sub business (unidad de negocio)
        strLinea = strLinea & separador                                         'pos 14 sub business (el cliente pidio quitarlo)
        strLinea = strLinea & separador                                         'pos 15 band
        strLinea = strLinea & separador & obtenerEstructura(ternro, 4, Date)    'pos 16 cargo (asumimos puesto)
        
        If EsNulo(rsDatosEmp!empremu) Then
            strLinea = strLinea & separador & "0"                               'pos 17 cargo (asumimos puesto)
        Else
            strLinea = strLinea & separador & CStr(rsDatosEmp!empremu)
        End If
        strLinea = strLinea & separador & armarDireccion(ternro, separador)     'pos 19-28 direccion y telefono

        strLinea = strLinea & separador & rsDatosEmp!empemail                   'pos 29 email
    
        'Cuil
        strLinea = strLinea & separador & IIf(EsNulo(rsDatosEmp!Cuil), "", rsDatosEmp!Cuil) 'pos 29 cuil
        
        'DNI
        strLinea = strLinea & separador & IIf(EsNulo(rsDatosEmp!dni), "", rsDatosEmp!dni) 'pos 30 cuil
    
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El empleado no tiene cargado algunno de los siguientes datos: "
        Flog.writeline Espacios(Tabulador * 2) & "Empresa, nacionalidad, pais de nacimiento o estado civil."
    End If
    
    'buscamos el reporta A
'    StrSql = "SELECT empleg FROM empleado WHERE empleado.ternro= " & IIf(EsNulo(rsDatosEmp!empreporta), 0, rsDatosEmp!empreporta)
'    OpenRecordset StrSql, rsAux
'    If rsAux.EOF Then
'        Flog.writeline Espacios(Tabulador * 1) & "El empleado no posee reporta A."
'        strLinea = strLinea & separador & "N/A"
'    Else
'        strLinea = strLinea & separador & rsAux!empleg
'    End If
    
    
    GoTo datosOk
CE:
    strLinea = ""
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Flog.writeline Espacios(Tabulador * 0) & "Error al tratar de recuperar los datos del modelo 1000. "
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & "SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Exit Function
datosOk:
    
    expEmpleado = strLinea
    Set rsDatosEmp = Nothing
End Function

