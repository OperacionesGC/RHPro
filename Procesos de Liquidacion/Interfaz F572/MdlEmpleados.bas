Attribute VB_Name = "MdlEmpleados"
Option Explicit

Dim Ternro As String
Dim l_sql As String
Dim NroDom As Integer
Dim empleg As String
Dim terape As String
Dim terape2 As String
Dim ternom As String
Dim ternom2 As String
Dim terfecnac As String
Dim tersex As String
Dim teremail As String
Dim nacionalnro As String
Dim nacionalidad As String
Dim estcivnro As String
Dim empfecalta As String
Dim empremu As String
Dim calle As String
Dim callenro As String
Dim piso As String
Dim depto As String
Dim locnro As String
Dim provnro As String
Dim codigopostal As String
Dim telnro As String
Dim paisnro As String
Dim NroDoc As String
Dim nrocuil As String
Dim empresaEstrnro As String
Dim tipmotnro As String
Dim caunro As String
Dim Estrnro As String
Dim Mensaje As String
Dim Empnro As String

Public Sub buscarempleado(ByVal l_cuit As String)
Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim l_cuit_No_Guion
Dim I As Integer

l_cuit_No_Guion = Replace(l_cuit, "-", "")
StrSql = " SELECT ternom,nrodoc,empleg, empleado.ternro FROM ter_doc "
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = ter_doc.ternro "
StrSql = StrSql & " WHERE ter_doc.tidnro in (6,10) AND "
StrSql = StrSql & " (nrodoc = '" & l_cuit & "'"
StrSql = StrSql & " Or nrodoc = '" & l_cuit_No_Guion & "'" & ")"
OpenRecordset StrSql, rsConsult
cuit = ""
For I = 0 To 5
    Arr_NroTernro(I) = 0
Next
I = 0
If Not rsConsult.EOF Then
    cuit = rsConsult!NroDoc
    While Not rsConsult.EOF
        Arr_NroTernro(I) = rsConsult!Ternro
        rsConsult.MoveNext
        I = I + 1
    Wend
Else
    If (Len(l_cuit_No_Guion) < 13) And (InStr(1, "-", l_cuit) = 0) Then
        l_cuit = Left(l_cuit_No_Guion, 2) & "-" & Mid(l_cuit_No_Guion, 3, 8) & "-" & Right(l_cuit_No_Guion, 1)
        
        StrSql = " SELECT ternom,nrodoc,empleg, empleado.ternro FROM ter_doc "
        StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = ter_doc.ternro "
        StrSql = StrSql & " WHERE ter_doc.tidnro in (6,10) AND "
        StrSql = StrSql & " (nrodoc = '" & l_cuit & "'"
        StrSql = StrSql & " Or nrodoc = '" & l_cuit_No_Guion & "'" & ")"
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
            cuit = rsConsult!NroDoc
             While Not rsConsult.EOF
                Arr_NroTernro(I) = rsConsult!Ternro
                rsConsult.MoveNext
                I = I + 1
             Wend
        End If
    End If
End If
rsConsult.Close
End Sub

Public Sub Generar_Empleado(node As IXMLDOMNode)
    Dim error As Boolean
    error = False
    Dim b As Long
    Dim l_cuit
    On Error GoTo ErrorEmpleado
    l_cuit = ""
    For b = 0 To node.childNodes.length - 1
        Select Case CStr(node.childNodes(b).nodeName)
            Case "cuit"
                l_cuit = node.childNodes.Item(b).Text
                buscarempleado (l_cuit) 'Buscar Legajo
        End Select
    Next b
    Err.Clear
    
    On Error GoTo ErrorEmpleado
    Err.Clear
    
    Exit Sub
ErrorEmpleado:
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & error
    Flog.writeline Espacios(Tabulador * 0) & "Linea " & RegLeidos & " del archivo procesado"
    HuboError = True
    HuboErrorLocal = True
    Exit Sub
End Sub

Public Sub Generar_PeriodosDeduc(node As IXMLDOMNode, doc, b, I)
    Dim error As Boolean
    Dim nodes As IXMLDOMNodeList
    error = False
    On Error GoTo ErrorEmpleado
    Dim l_TotalHijosPeriodo
    Dim l_HijoNombre
    Dim j
    Dim l_tipoDeduccion
    l_tipoDeduccion = node.childNodes.Item(b).Attributes.getNamedItem("tipo").Text
    l_TotalHijosPeriodo = node.childNodes.Item(b).childNodes.Item(I).childNodes.length
    If l_tipoDeduccion <> 10 Then
        For j = 0 To l_TotalHijosPeriodo - 1
          Deducciones(b).periodosdeduc(j).periodo_mesDesde = node.childNodes(b).childNodes(I).childNodes(j).Attributes.getNamedItem("mesDesde").Text
          Deducciones(b).periodosdeduc(j).periodo_meshasta = node.childNodes(b).childNodes(I).childNodes(j).Attributes.getNamedItem("mesHasta").Text
          Deducciones(b).periodosdeduc(j).periodo_montoMensual = node.childNodes(b).childNodes(I).childNodes(j).Attributes.getNamedItem("montoMensual").Text
          ReDim Preserve Deducciones(b).periodosdeduc(j + 1)
        Next
    Else
        For j = 0 To l_TotalHijosPeriodo - 1
            Deducciones(b).periodosdeduc(j).periodo_mesDesde = node.childNodes(b).childNodes(I).childNodes(j).Attributes.getNamedItem("mesDesde").Text
            Deducciones(b).periodosdeduc(j).periodo_meshasta = node.childNodes(b).childNodes(I).childNodes(j).Attributes.getNamedItem("mesHasta").Text
            ReDim Preserve Deducciones(b).periodosdeduc(j + 1)
        Next
    End If
    Err.Clear
    On Error GoTo ErrorEmpleado
    Err.Clear
   
    Exit Sub
ErrorEmpleado:
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & error
    Flog.writeline Espacios(Tabulador * 0) & "Linea " & RegLeidos & " del archivo procesado"
    If rs.State = adStateOpen Then rs.Close
    HuboError = True
    HuboErrorLocal = True
    Exit Sub
End Sub

Public Sub Generar_DetallesDeduc(node As IXMLDOMNode, doc, b, I)
    Dim error As Boolean
    Dim nodes As IXMLDOMNodeList
    error = False
    On Error GoTo ErrorEmpleado
    Dim l_TotalHijosPeriodo
    Dim l_HijoNombre
    Dim j
    
    'l_TotalHijosPeriodo = node.childNodes.Item(b).childNodes.Item(I).childNodes.length
    'For j = 0 To l_TotalHijosPeriodo - 2
      'Deducciones(b).detallesdeduc(j).detalle_nombre = node.childNodes(b).childNodes(I).childNodes(j).Attributes.getNamedItem("nombre").Text
    '  Deducciones(b).detallesdeduc(j).detalle_valor = node.childNodes(b).childNodes(I).childNodes(j).Attributes.getNamedItem("valor").Text
    '   ReDim Preserve Deducciones(b).detallesdeduc(j + 1)
    'Next

        '--------------------------------------MDF
        If node.childNodes(b).childNodes(I).childNodes(0).Attributes.getNamedItem("nombre").Text = "fechaAporte" Then
          Deducciones(b).Mes_periodo = node.childNodes(b).childNodes(I).childNodes(0).Attributes.getNamedItem("valor").Text
          Deducciones(b).Mes_periodo = Month(CDate(Deducciones(b).Mes_periodo))
        Else
         Deducciones(b).Mes_periodo = ""
        End If
        Deducciones(b).Mes = node.childNodes(b).childNodes(I).childNodes(1).Attributes.getNamedItem("valor").Text
        '-------------------------------------MDF
       
    Err.Clear
    On Error GoTo ErrorEmpleado
    Err.Clear
   
    Exit Sub
ErrorEmpleado:
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & error
    Flog.writeline Espacios(Tabulador * 0) & "Linea " & RegLeidos & " del archivo procesado"
    If rs.State = adStateOpen Then rs.Close
    HuboError = True
    HuboErrorLocal = True
    Exit Sub
End Sub

Public Sub Generar_DetallesRetPagos(node As IXMLDOMNode, doc, b, I)
    Dim error As Boolean
    Dim nodes As IXMLDOMNodeList
    error = False
    On Error GoTo ErrorEmpleado
    Dim l_TotalHijosPeriodo
    Dim l_HijoNombre
    Dim j
    
    l_TotalHijosPeriodo = node.childNodes.Item(b).childNodes.Item(I).childNodes.length
    For j = 0 To l_TotalHijosPeriodo - 1
      retPerPagos(b).detallesretpagos(j).detalle_nombre = node.childNodes(b).childNodes(I).childNodes(j).Attributes.getNamedItem("nombre").Text
      retPerPagos(b).detallesretpagos(j).detalle_valor = node.childNodes(b).childNodes(I).childNodes(j).Attributes.getNamedItem("valor").Text
      ReDim Preserve retPerPagos(b).detallesretpagos(j + 1)
    Next j

    Err.Clear
    On Error GoTo ErrorEmpleado
    Err.Clear
   
    Exit Sub
ErrorEmpleado:
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & error
    Flog.writeline Espacios(Tabulador * 0) & "Linea " & RegLeidos & " del archivo procesado"
    If rs.State = adStateOpen Then rs.Close
    HuboError = True
    HuboErrorLocal = True
    Exit Sub
End Sub

Public Sub Generar_PeriodosRetPagos(node As IXMLDOMNode, doc, b, I)
    Dim error As Boolean
    Dim nodes As IXMLDOMNodeList
    error = False
    On Error GoTo ErrorEmpleado
    Dim l_TotalHijosPeriodo
    Dim l_HijoNombre
    Dim j
    
    l_TotalHijosPeriodo = node.childNodes.Item(b).childNodes.Item(I).childNodes.length
    For j = 0 To l_TotalHijosPeriodo - 1
      retPerPagos(b).periodosretpagos(j).periodo_mesDesde = node.childNodes(b).childNodes(I).childNodes(j).Attributes.getNamedItem("mesDesde").Text
      retPerPagos(b).periodosretpagos(j).periodo_meshasta = node.childNodes(b).childNodes(I).childNodes(j).Attributes.getNamedItem("mesHasta").Text
      retPerPagos(b).periodosretpagos(j).periodo_montoMensual = node.childNodes(b).childNodes(I).childNodes(j).Attributes.getNamedItem("montoMensual").Text
      ReDim Preserve retPerPagos(b).periodosretpagos(j + 1)
    Next j

    Err.Clear
    On Error GoTo ErrorEmpleado
    Err.Clear
   
    Exit Sub
ErrorEmpleado:
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & error
    Flog.writeline Espacios(Tabulador * 0) & "Linea " & RegLeidos & " del archivo procesado"
    If rs.State = adStateOpen Then rs.Close
    HuboError = True
    HuboErrorLocal = True
    Exit Sub
End Sub

Public Sub ingresosAportes(node, doc, b, I, j)
    Dim error As Boolean
    Dim nodes As IXMLDOMNodeList
    error = False
    On Error GoTo ErrorEmpleado
    Dim l_TotalHijosPeriodo
    Dim l_HijoNombre
    Dim l_ganliq_mes
    
      GanLiq(b).ingAp(j).obrasoc = node.childNodes.Item(b).childNodes.Item(I).childNodes(j).childNodes(0).Text
      GanLiq(b).ingAp(j).segsocial = node.childNodes.Item(b).childNodes.Item(I).childNodes(j).childNodes(1).Text
      GanLiq(b).ingAp(j).sind = node.childNodes.Item(b).childNodes.Item(I).childNodes(j).childNodes(2).Text
      GanLiq(b).ingAp(j).ganbrut = node.childNodes.Item(b).childNodes.Item(I).childNodes(j).childNodes(3).Text
      GanLiq(b).ingAp(j).retgan = node.childNodes.Item(b).childNodes.Item(I).childNodes(j).childNodes(4).Text
      GanLiq(b).ingAp(j).retribNoHab = node.childNodes.Item(b).childNodes.Item(I).childNodes(j).childNodes(5).Text
      GanLiq(b).ingAp(j).ajuste = node.childNodes.Item(b).childNodes.Item(I).childNodes(j).childNodes(6).Text
      GanLiq(b).ingAp(j).Mes = node.childNodes(b).childNodes(I).childNodes(j).Attributes.getNamedItem("mes").Text
      
      ReDim Preserve GanLiq(b).ingAp(j + 1)

    Err.Clear
    On Error GoTo ErrorEmpleado
    Err.Clear
   
    Exit Sub
ErrorEmpleado:
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & error
    Flog.writeline Espacios(Tabulador * 0) & "Linea " & RegLeidos & " del archivo procesado"
    If rs.State = adStateOpen Then rs.Close
    HuboError = True
    HuboErrorLocal = True
    Exit Sub

End Sub

Public Sub Grabar_GanLiq(l_cuit, l_denominacion, LimiteSuperior)
Dim I

'For I = 0 To LimiteSuperior
'    GanLiq(I).cuit = l_cuit
'    GanLiq(I).Denominacion = l_denominacion
'Next I
GanLiq(LimiteSuperior).cuit = l_cuit
GanLiq(LimiteSuperior).Denominacion = l_denominacion
End Sub
Public Sub Grabar_Deducciones(l_deduccion, l_tipoDoc, l_Nrodoc, l_denominacion, l_descBasica, l_descAdicional, l_montoTotal, l_periodo, l_detalles, LimiteSuperior)

Deducciones(LimiteSuperior).tipo = l_deduccion
Deducciones(LimiteSuperior).TipoDoc = l_tipoDoc
Deducciones(LimiteSuperior).NroDoc = l_Nrodoc
Deducciones(LimiteSuperior).Denominacion = l_denominacion
Deducciones(LimiteSuperior).DescBasica = l_descBasica
Deducciones(LimiteSuperior).DescAdicional = l_descAdicional
Deducciones(LimiteSuperior).MontoTotal = l_montoTotal
'Deducciones(UBound(Deducciones)).detalles = l_detalles
End Sub
'FB - Se creo este sub para grabar las deducciones de tipo 10
Public Sub Grabar_DeduccionesTipo10(l_deduccion, l_montoTotal, l_periodo, l_detalles, LimiteSuperior)
Deducciones(LimiteSuperior).tipo = l_deduccion
Deducciones(LimiteSuperior).periodosdeduc(0).periodo_anio = l_periodo
Deducciones(LimiteSuperior).MontoTotal = l_montoTotal
Deducciones(UBound(Deducciones)).detallesdeduc(1).detalle_valor = l_detalles
End Sub

Public Sub Grabar_CargasFamilia(l_tipoDoc, l_Nrodoc, l_apellido, l_nombre, l_fechaNac, l_mesDesde, l_mesHasta, l_parentesco, LimiteSuperior)

CargaFamilia(LimiteSuperior).TipoDoc = l_tipoDoc
CargaFamilia(LimiteSuperior).NroDoc = l_Nrodoc
CargaFamilia(LimiteSuperior).Apellido = l_apellido
CargaFamilia(LimiteSuperior).Nombre = l_nombre
CargaFamilia(LimiteSuperior).FechaNac = l_fechaNac
CargaFamilia(LimiteSuperior).MesDesde = l_mesDesde
CargaFamilia(LimiteSuperior).MesHasta = l_mesHasta
CargaFamilia(LimiteSuperior).Parentesco = l_parentesco
End Sub

Public Sub Grabar_Retenciones(l_retPerPago, l_tipoDoc, l_Nrodoc, l_denominacion, l_descBasica, l_descAdicional, l_montoTotal, l_periodo, l_detalles, LimiteSuperior)

retPerPagos(LimiteSuperior).tipo = l_retPerPago
retPerPagos(LimiteSuperior).TipoDoc = l_tipoDoc
retPerPagos(LimiteSuperior).NroDoc = l_Nrodoc
retPerPagos(LimiteSuperior).Denominacion = l_denominacion
retPerPagos(LimiteSuperior).DescBasica = l_descBasica
retPerPagos(LimiteSuperior).DescAdicional = l_descAdicional
retPerPagos(LimiteSuperior).MontoTotal = l_montoTotal
End Sub

Public Sub Generar_Deducciones(node As IXMLDOMNode, doc)
    Dim error As Boolean
    Dim nodes As IXMLDOMNodeList
    error = False
    Dim b As Long
    On Error GoTo ErrorEmpleado
    Dim l_TotalHijosDesduccion
    Dim l_HijoNombre
    Dim I
    Dim l_tipoDoc
    Dim l_mes
    Dim l_Nrodoc
    Dim l_denominacion
    Dim l_descBasica
    Dim l_descAdicional
    Dim l_montoTotal As String
    Dim l_periodo
    Dim l_detalles
    Dim l_deduccion
    Dim l_TotalHijosPeriodo
    Dim j
    Dim l_TotalHijosDeduc
    
    ReDim Preserve Deducciones(1)
    LimiteSuperior = 0
    l_tipoDoc = "0"
    l_denominacion = ""
    l_descBasica = ""
    l_descAdicional = "0"
    l_montoTotal = "0"
    l_periodo = "0"
    For b = 0 To (node.childNodes.length - 1)
       ReDim Preserve Deducciones(b).periodosdeduc(1)
       ReDim Preserve Deducciones(b).detallesdeduc(1)
       l_TotalHijosDesduccion = node.childNodes.Item(b).childNodes.length
       l_deduccion = node.childNodes.Item(b).Attributes.getNamedItem("tipo").Text
       For I = 0 To l_TotalHijosDesduccion - 1
        l_HijoNombre = node.childNodes(b).childNodes(I).nodeName
            Select Case l_HijoNombre
                Case "tipoDoc"
                    l_tipoDoc = node.childNodes.Item(b).childNodes(I).Text
                Case "nroDoc"
                    l_Nrodoc = node.childNodes.Item(b).childNodes(I).Text
                Case "denominacion"
                    l_denominacion = node.childNodes.Item(b).childNodes(I).Text
                Case "descBasica"
                    l_descBasica = node.childNodes.Item(b).childNodes(I).Text
                Case "descAdicional"
                    l_descAdicional = node.childNodes.Item(b).childNodes(I).Text
                Case "montoTotal"
                    l_montoTotal = node.childNodes.Item(b).childNodes(I).Text
                Case "periodos"
                    'FB - Se obtiene el periodo para el tipo de deduccion 10
                    If l_deduccion = 10 Then
                        Set nodes = doc.selectNodes("presentacion/periodo")
                        l_periodo = nodes.Item(0).Text
                    Else
                        Call Generar_PeriodosDeduc(node, doc, b, I)
                    End If
                    'FB
                Case "detalles"
                    'FB - Si la deduccion es de tipo 10, se obtienen los detalles para esta deduccion
                    If l_deduccion = 10 Then
                        Set nodes = doc.selectNodes("presentacion/deducciones")
                        l_TotalHijosDeduc = node.childNodes.Item(b).childNodes.Item(I).childNodes.length
                        For j = 0 To l_TotalHijosDeduc - 1
                            If node.childNodes(b).childNodes(I).childNodes(j).Attributes.getNamedItem("nombre").Text = "desc" Then
                                Deducciones(b).detallesdeduc(0).detalle_nombre = node.childNodes(b).childNodes(I).childNodes(j).Attributes.getNamedItem("nombre").Text
                                Deducciones(b).detallesdeduc(1).detalle_valor = node.childNodes(b).childNodes(I).childNodes(j).Attributes.getNamedItem("valor").Text
                                l_detalles = Deducciones(b).detallesdeduc(1).detalle_valor
                            End If
                        Next
                    'FB
                    Else
                        Call Generar_DetallesDeduc(node, doc, b, I)
                    End If
                    'l_mes = node.childNodes(b).childNodes(I).childNodes(j).Attributes.getNamedItem("valor").Text
            End Select
        Next
        'FB - Si la deduccion es de tipo 10, se creo la funcion para insertar las deducciones de tipo 10
        If l_deduccion = 10 Then
            Call Grabar_DeduccionesTipo10(l_deduccion, l_montoTotal, l_periodo, l_detalles, LimiteSuperior)
        Else
            Call Grabar_Deducciones(l_deduccion, l_tipoDoc, l_Nrodoc, l_denominacion, l_descBasica, l_descAdicional, l_montoTotal, l_periodo, l_detalles, LimiteSuperior)
        End If
        LimiteSuperior = LimiteSuperior + 1
        ReDim Preserve Deducciones(b + 1)
    Next b
    Err.Clear
    On Error GoTo ErrorEmpleado
    Err.Clear
   
    Exit Sub
ErrorEmpleado:
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & error
    Flog.writeline Espacios(Tabulador * 0) & "Linea " & RegLeidos & " del archivo procesado"
    If rs.State = adStateOpen Then rs.Close
    HuboError = True
    HuboErrorLocal = True
    Exit Sub
End Sub
Public Sub Generar_CargasFamilia(node As IXMLDOMNode, doc)
    Dim error As Boolean
    Dim nodes As IXMLDOMNodeList
    error = False
    Dim b As Long
    On Error GoTo ErrorEmpleado
    Dim l_TotalHijosCargaFam
    Dim l_HijoNombre
    Dim I
    Dim l_tipoDocFam
    Dim l_NrodocFam
    Dim l_apellidoFam
    Dim l_nombreFam
    Dim l_fechaNacFam
    Dim l_mesDesdeFam
    Dim l_mesHastaFam
    Dim l_parentesco
    ReDim Preserve CargaFamilia(1)
    
    LimiteSuperior = 0
    l_tipoDocFam = "0"
    l_apellidoFam = ""
    l_nombreFam = ""
    l_fechaNacFam = ""
    l_mesDesdeFam = 0
    l_mesHastaFam = 0
    l_parentesco = 0
    
    For b = 0 To (node.childNodes.length - 1)
       l_TotalHijosCargaFam = node.childNodes.Item(b).childNodes.length
       For I = 0 To l_TotalHijosCargaFam - 1
        l_HijoNombre = node.childNodes(b).childNodes(I).nodeName
            Select Case l_HijoNombre
                Case "tipoDoc"
                    l_tipoDocFam = node.childNodes.Item(b).childNodes(I).Text
                Case "nroDoc"
                    l_NrodocFam = node.childNodes.Item(b).childNodes(I).Text
                Case "apellido"
                    l_apellidoFam = node.childNodes.Item(b).childNodes(I).Text
                Case "nombre"
                    l_nombreFam = node.childNodes.Item(b).childNodes(I).Text
                Case "fechaNac"
                    l_fechaNacFam = node.childNodes.Item(b).childNodes(I).Text
                Case "mesDesde"
                    l_mesDesdeFam = node.childNodes.Item(b).childNodes(I).Text
                Case "mesHasta"
                    l_mesHastaFam = node.childNodes.Item(b).childNodes(I).Text
                Case "parentesco"
                    l_parentesco = node.childNodes.Item(b).childNodes(I).Text
            End Select
        Next
        Call Grabar_CargasFamilia(l_tipoDocFam, l_NrodocFam, l_apellidoFam, l_nombreFam, l_fechaNacFam, l_mesDesdeFam, l_mesHastaFam, l_parentesco, LimiteSuperior)
        LimiteSuperior = LimiteSuperior + 1
        ReDim Preserve CargaFamilia(b + 1)
    Next b
    Err.Clear
    On Error GoTo ErrorEmpleado
    Err.Clear
   
    Exit Sub
ErrorEmpleado:
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & error
    Flog.writeline Espacios(Tabulador * 0) & "Linea " & RegLeidos & " del archivo procesado"
    If rs.State = adStateOpen Then rs.Close
    HuboError = True
    HuboErrorLocal = True
    Exit Sub
End Sub

Public Sub Generar_GanLiq(node As IXMLDOMNode, doc)
    Dim error As Boolean
    Dim nodes As IXMLDOMNodeList
    error = False
    Dim b As Long
    On Error GoTo ErrorEmpleado
    Dim l_TotalHijosGanliq
    Dim l_HijoNombre
    Dim I, j
    Dim l_cuit
    Dim l_denominacion
    Dim l_ganliq_mes
    Dim l_TotalHijosAp
    ReDim Preserve GanLiq(1)
    
    LimiteSuperior = 0
    l_cuit = ""
    l_denominacion = ""
    l_TotalHijosAp = "0"
    For b = 0 To (node.childNodes.length - 1)
       ReDim Preserve GanLiq(b).ingAp(1)
       l_TotalHijosGanliq = node.childNodes.Item(b).childNodes.length
       For I = 0 To l_TotalHijosGanliq - 1
        l_HijoNombre = node.childNodes(b).childNodes(I).nodeName
            Select Case l_HijoNombre
                Case "cuit"
                    l_cuit = node.childNodes.Item(b).childNodes(I).Text
                Case "denominacion"
                    l_denominacion = node.childNodes.Item(b).childNodes(I).Text
                Case "ingresosAportes"
                    l_TotalHijosAp = node.childNodes.Item(b).childNodes(I).childNodes.length
                    For j = 0 To l_TotalHijosAp - 1
                          'l_ganliq_mes = node.childNodes(b).childNodes(i).childNodes(j).Attributes.getNamedItem("mes").Text
                    Call ingresosAportes(node, doc, b, I, j)
                    Next
                    'FBDenominacion = Deducciones(I).Denominacion
            End Select
        Next
        Call Grabar_GanLiq(l_cuit, l_denominacion, LimiteSuperior)
        ReDim Preserve GanLiq(b + 1)
        LimiteSuperior = LimiteSuperior + 1
    Next b
    Err.Clear
    On Error GoTo ErrorEmpleado
    Err.Clear
   
    Exit Sub
ErrorEmpleado:
    Flog.writeline Espacios(Tabulador * 0) & "Error10: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & error
    Flog.writeline Espacios(Tabulador * 0) & "Linea " & RegLeidos & " del archivo procesado"
    If rs.State = adStateOpen Then rs.Close
    HuboError = True
    HuboErrorLocal = True
    Exit Sub
End Sub

Public Sub Generar_Retenciones(node As IXMLDOMNode, doc)
    Dim error As Boolean
    Dim nodes As IXMLDOMNodeList
    error = False
    Dim b As Long
    On Error GoTo ErrorEmpleado
    Dim l_TotalHijosRetencion
    Dim l_HijoNombre
    Dim I
    Dim l_tipoDoc
    Dim l_Nrodoc
    Dim l_denominacion
    Dim l_descBasica
    Dim l_descAdicional
    Dim l_montoTotal As String
    Dim l_detalles
    Dim l_periodo
    Dim l_retPerPago
    ReDim Preserve retPerPagos(1)
    
    LimiteSuperior = 0
    l_tipoDoc = "0"
    l_denominacion = ""
    l_descBasica = ""
     l_descAdicional = ""
     l_montoTotal = "0"
    For b = 0 To (node.childNodes.length - 1)
       l_TotalHijosRetencion = node.childNodes.Item(b).childNodes.length
       l_retPerPago = node.childNodes.Item(b).Attributes.getNamedItem("tipo").Text
       ReDim Preserve retPerPagos(b).detallesretpagos(1)
       ReDim Preserve retPerPagos(b).periodosretpagos(1)
       For I = 0 To l_TotalHijosRetencion - 1
        l_HijoNombre = node.childNodes(b).childNodes(I).nodeName
            Select Case l_HijoNombre
                Case "tipoDoc"
                    l_tipoDoc = node.childNodes.Item(b).childNodes(I).Text
                Case "nroDoc"
                    l_Nrodoc = node.childNodes.Item(b).childNodes(I).Text
                Case "denominacion"
                    l_denominacion = node.childNodes.Item(b).childNodes(I).Text
                Case "descBasica"
                    l_descBasica = node.childNodes.Item(b).childNodes(I).Text
                Case "descAdicional"
                    l_descAdicional = node.childNodes.Item(b).childNodes(I).Text
                Case "montoTotal"
                    l_montoTotal = node.childNodes.Item(b).childNodes(I).Text
                Case "periodos"
                    Call Generar_PeriodosRetPagos(node, doc, b, I)
                Case "detalles"
                    Call Generar_DetallesRetPagos(node, doc, b, I)
            End Select
        Next
        Call Grabar_Retenciones(l_retPerPago, l_tipoDoc, l_Nrodoc, l_denominacion, l_descBasica, l_descAdicional, l_montoTotal, l_periodo, l_detalles, LimiteSuperior)
        LimiteSuperior = LimiteSuperior + 1
        ReDim Preserve retPerPagos(b + 1)
    Next b
    Err.Clear
    On Error GoTo ErrorEmpleado
    Err.Clear
   
    Exit Sub
ErrorEmpleado:
    Flog.writeline Espacios(Tabulador * 0) & "Error11: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & error
    Flog.writeline Espacios(Tabulador * 0) & "Linea " & RegLeidos & " del archivo procesado"
    If rs.State = adStateOpen Then rs.Close
    HuboError = True
    HuboErrorLocal = True
    Exit Sub
End Sub
