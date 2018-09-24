Attribute VB_Name = "MdlModelos"
Option Explicit

Public Type TR_CBU_Bloque_1 'longitud total = 8 posiciones
    '   Codigo entidad:  "011"  - long 3
        Codigo_Entidad As String
    '   Codigo Sucursal: "BCRA" - long 4
        Codigo_Sucursal As String
    '   Digito Verif. Bloque 1  - long 1
        Digito_Verificador As String
End Type

Public Type TR_CBU_Bloque_2 'longitud total = 14 posiciones
    '   Tipo de Cuenta:         - long 1  (2 = CC, 3 = CA y 4 = CCE)
        Cuenta_Tipo As String
    '   Moneda de la cuenta:    - long 1  (0 = Pesos, 1 = Dolares y 3 = Lecop)
        Moneda As String
    '   Nro de la cuenta        - long 11
        Cuenta_Nro As String
    '   Digito Verif. Bloque 2  - long 1
        Digito_Verificador As String
End Type

Public Type TR_CBU 'longitud total = 22 posiciones
    Bloque1 As String 'longitud total = 8 posiciones
    Bloque2 As String 'longitud total = 14 posiciones
End Type

Public Type TR_Datos_Bancarios 'longitud total = 130 posiciones
    Proceso As String               'String   long 2  - Valor Fijo "AH"
    Servicio As String              'String   long 4  -
    Sucursal As String              'Numerico long 4  - Valor Fijo 0002
    Legajo As String                'Numerico long 20 -
    Moneda As String                'String   long 1  - Valor Fijo "P"
    Titularidad As String           'String   long 2  - Valor Fijo "SF"
    CBU As String                   'Numerico long 22 -
                                    'Bloque 1
                                    '       Codigo entidad:  "011"  - long 3
                                    '       Codigo Sucursal: "BCRA" - long 4
                                    '       Digito Verif. Bloque 1  - long 1
                                    'Bloque 2
                                    '       Tipo de Cuenta:         - long 1  (2 = CC, 3 = CA y 4 = CCE)
                                    '       Moneda de la cuenta:    - long 1  (0 = Pesos, 1 = Dolares y 3 = Lecop)
                                    '       Nro de la cuenta        - long 11
                                    '       Digito Verif. Bloque 2  - long 1
    Cuenta_Electronica As String    'Numerico long 19 - (nro de tarjeta de debito)
    Tarjeta_1er_Titular As String   'Numerico long 19 -
    Tarjeta_2do_Titular As String   'Numerico long 19 -
    Doc_Tipo As String              'String   long 2  -
    Doc_Nro As String               'Numerico long 11 -
    Filler As String                'String   long 5  -
End Type



Global NroModelo As Long
Global DescripcionModelo As String
Global Primera_Vez As Boolean
Global Banco As Long
'Variables globales levantadas como parametros del proceso (batch_proceso.bprcparam)
Global PisaNovedad As Boolean
Global TikPedNro As Long
Global NombreArchivo As String
Global acunro As Long 'se usa en el modelo 216 de Citrusvil y se carga por confrep


Public Sub Insertar_Linea_Segun_Modelo(ByVal Linea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento llamador de acurdo al modelo
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
MyBeginTrans
    Select Case NroModelo
    Case 112: 'Migracion de Estructuras
        Call LineaModelo_112(Linea)
    Case 211: 'Novedades
        Call LineaModelo_211(Linea)
'    Case 212: 'GTI - Mega Alarmas
'        Call LineaModelo_212(Linea)
'    Case 213: 'GTI - Acumulado Diario
'        Call LineaModelo_213(Linea)
    Case 214: 'Tickets
        Call LineaModelo_214(Linea)
    Case 215: 'Acumuladores de Agencia
        Call LineaModelo_215(Linea)
    Case 216: 'Acumuladores de Agencia para Citrusvil
        Call LineaModelo_216(Linea)
    Case 217: 'Vales
        Call LineaModelo_217(Linea)
    Case 218: 'Migracion de Novedades
        Call LineaModelo_218(Linea)
    Case 219: 'Libre
        Call LineaModelo_219(Linea)
    Case 220: 'Libre
        Call LineaModelo_220(Linea)
    Case 221: 'Libre
        Call LineaModelo_221(Linea)
    Case 222: 'DesmenFamiliar
        Call LineaModelo_222
    Case 223: '
        Call LineaModelo_223(Linea)
    Case 224: '
        Call LineaModelo_224(Linea)
    Case 225: '
        Call LineaModelo_225(Linea)
    Case 226: '
'        Call LineaModelo_226(Linea)
    Case 227: '
        Call LineaModelo_227(Linea)
    Case 228: '
        'Call LineaModelo_228(Linea)
        'Modelo reservado para el reporte de Declaracion Jurada
    Case 229: 'Prestamos
        Call LineaModelo_229(Linea)
    Case 230: '
        Call LineaModelo_230(Linea)
    Case 231: 'Interface Cta Banco Nacion
        Call LineaModelo_231(Linea)
    End Select
MyCommitTrans
End Sub


Public Sub LineaModelo_211(ByVal strLinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta Novedad segun formato
' Autor      : FGZ
'              El formato es:
'                   Formato 1
'                       Legajo; conccod; tpanro; monto
'                   ó
'                   Formato 2.1
'                       Legajo; conccod; tpanro; monto; FechaDesde; FechaHasta
'                   Formato 2.2
'                       Legajo; conccod; tpanro; monto; FechaDesde
'                   ó
'                   Formato 3
'                       Legajo; conccod; tpanro; monto; MarcaRetroactividad;PeriodoDesde(pliqnro); PeriodoHasta(pliqnro)
' Fecha      : 29/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
    
Dim Tercero As Long
Dim NroLegajo As Long

Dim concnro As Long
Dim Conccod As Long

Dim tpanro As Long
Dim Monto As Single
'Dim FechaDesde As Date
'Dim FechaHasta As Date
Dim FechaDesde As String
Dim FechaHasta As String

Dim PeriodoDesde As Long
Dim PeriodoHasta As Long
Dim TieneVigencia As Boolean
Dim EsRetroactivo As Boolean

Dim aux As String

Dim rs_Empleado As New ADODB.Recordset
Dim rs_Concepto As New ADODB.Recordset
Dim rs_TipoPar As New ADODB.Recordset
Dim rs_NovEmp As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset


' El formato es:
' Formato 1
' Legajo; conccod; tpanro; monto
' ó
' Formato 2
' Legajo; conccod; tpanro; monto; FechaDesde; FechaHasta
' ó
' Formato 3
' Legajo; conccod; tpanro; monto; MarcaRetroactividad;PeriodoDesde(pliqnro); PeriodoHasta(pliqnro)
    
    TieneVigencia = False
    EsRetroactivo = False

    'Nro de Legajo
    pos1 = 1
    pos2 = InStr(pos1, strLinea, Separador)
    If IsNumeric(Mid$(strLinea, pos1, pos2 - pos1)) Then
        NroLegajo = Mid$(strLinea, pos1, pos2 - pos1)
    Else
        InsertaError 1, 8
        Exit Sub
    End If
    
    'Concepto
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    Conccod = Mid(strLinea, pos1, pos2 - pos1)

    'Tipo de Parametro
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    tpanro = Mid(strLinea, pos1, pos2 - pos1)

    'Monto
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    If pos2 = 0 Then
        pos2 = Len(strLinea)
        Monto = Mid(strLinea, pos1, pos2)
        Monto = CSng(Replace(CStr(Monto), SeparadorDecimal, "."))
    Else
        Monto = Mid(strLinea, pos1, pos2 - pos1)
        Monto = CSng(Replace(CStr(Monto), SeparadorDecimal, "."))
        
        'Puede veniar Fecha Desde; fecha Hasta ó Retroactivo, Periodo desde , Periodo Hasta
        pos1 = pos2 + 1
        pos2 = InStr(pos1 + 1, strLinea, Separador)
        If pos2 > 0 Then
            aux = Mid(strLinea, pos1, pos2 - pos1)
            If IsDate(aux) Then
                'Fecha desde
                FechaDesde = Mid(strLinea, pos1, pos2 - pos1)
            
                'Fecha Hasta
                pos1 = pos2 + 1
                pos2 = Len(strLinea)
                FechaHasta = Mid(strLinea, pos1, pos2)
                
                TieneVigencia = True
            Else
                If UCase(aux) = "SI" Then
                    EsRetroactivo = True
                Else
                    EsRetroactivo = False
                End If
                
                'Periodo desde
                pos1 = pos2 + 1
                pos2 = InStr(pos1 + 1, strLinea, Separador)
                PeriodoDesde = Mid(strLinea, pos1, pos2 - pos1)
                
                'Periodo hasta
                pos1 = pos2 + 1
                pos2 = Len(strLinea)
                PeriodoHasta = Mid(strLinea, pos1, pos2)
            End If
        Else
            'Viene Vigencia con fecha desde y sin fecha hasta
            pos2 = Len(strLinea)
            FechaDesde = Mid(strLinea, pos1, pos2)
            TieneVigencia = True
            FechaHasta = ""
        End If
    End If

' ====================================================================
'   Validar los parametros Levantados

'Que exista el legajo
StrSql = "SELECT * FROM empleado where empleg = " & NroLegajo
OpenRecordset StrSql, rs_Empleado
If rs_Empleado.EOF Then
    Flog.Writeline "No se encontro el legajo " & NroLegajo
    InsertaError 1, 8
    Exit Sub
Else
    Tercero = rs_Empleado!ternro
End If

'Que exista el concepto
StrSql = "SELECT * FROM concepto WHERE conccod = " & Conccod
OpenRecordset StrSql, rs_Concepto
If rs_Concepto.EOF Then
    Flog.Writeline "No se encontro el Concepto " & Conccod
    InsertaError 2, 10
    Exit Sub
Else
    concnro = rs_Concepto!concnro
End If

'Que exista el tipo de Parametro
StrSql = "SELECT * FROM tipopar WHERE tpanro = " & tpanro
OpenRecordset StrSql, rs_TipoPar

If rs_TipoPar.EOF Then
    Flog.Writeline "No se encontro el Tipo de Parametro " & tpanro
    InsertaError 3, 11
    Exit Sub
End If


If EsRetroactivo Then
    'Chequeo que los periodos sean validos
    'Chequeo Periodo Desde
    StrSql = "SELECT * FROM periodo WHERE pliqnro = " & PeriodoDesde
    If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
    OpenRecordset StrSql, rs_Periodo
    
    If rs_Periodo.EOF Then
        Flog.Writeline "Periodo Desde Invalido " & PeriodoDesde
        InsertaError 6, 36
        Exit Sub
    End If
    
    'Chequeo Periodo Hasta
    StrSql = "SELECT * FROM periodo WHERE pliqnro = " & PeriodoHasta
    If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
    OpenRecordset StrSql, rs_Periodo
    
    If rs_Periodo.EOF Then
        Flog.Writeline "Periodo Hasta Invalido " & PeriodoHasta
        InsertaError 7, 36
        Exit Sub
    End If
End If

'=============================================================
'Busco si existe la Novedad
If Not TieneVigencia Then
    StrSql = "SELECT * FROM novemp WHERE "
    StrSql = StrSql & " concnro = " & concnro
    StrSql = StrSql & " AND tpanro = " & tpanro
    StrSql = StrSql & " AND empleado = " & Tercero
    StrSql = StrSql & " AND (nevigencia = -1 OR nevigencia = 0) "
    StrSql = StrSql & " ORDER BY nevigencia "
    If rs_NovEmp.State = adStateOpen Then rs_NovEmp.Close
    OpenRecordset StrSql, rs_NovEmp

    If Not rs_NovEmp.EOF Then
        If CBool(rs_NovEmp!nevigencia) Then
            Flog.Writeline "No se puede insertar la novedad poqrue ya existe una con Vigencia"
            InsertaError 1, 94
            Exit Sub
        Else
            'Existe una novedad pero sin vigencia ==> Actualizo
            If PisaNovedad Then 'Actualizo la Novedad
                If Not EsRetroactivo Then
                    StrSql = "UPDATE novemp SET nevalor = " & Monto & _
                             " , neretro = 0 " & _
                             " WHERE concnro = " & concnro & _
                             " AND tpanro = " & tpanro & _
                             " AND empleado = " & Tercero
                Else
                    StrSql = "UPDATE novemp SET nevalor = " & Monto & _
                             " , neretro = -1 " & _
                             " , nepliqdesde =  " & PeriodoDesde & _
                             " , nepliqhasta =  " & PeriodoHasta & _
                             " WHERE concnro = " & concnro & _
                             " AND tpanro = " & tpanro & _
                             " AND empleado = " & Tercero
                End If
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.Writeline "Novedad Actualizada "
            Else
                Flog.Writeline "No se insertó la novedad. Ya existe y no se pisa."
            End If
        End If
    Else
        'Inserto
        If Not EsRetroactivo Then
            StrSql = "INSERT INTO novemp (" & _
                     "empleado,concnro,tpanro,nevalor,nevigencia" & _
                     ") VALUES (" & Tercero & _
                     "," & concnro & _
                     "," & tpanro & _
                     "," & Monto & _
                     ",0" & _
                     " )"
        Else
            StrSql = "INSERT INTO novemp (" & _
                     "empleado,concnro,tpanro,nevalor,nevigencia,neretro,nepliqdesde,nepliqhasta " & _
                     ") VALUES (" & Tercero & _
                     "," & concnro & _
                     "," & tpanro & _
                     "," & Monto & _
                     ",0" & _
                     "," & CInt(EsRetroactivo) & _
                     "," & PeriodoDesde & _
                     "," & PeriodoHasta & _
                     " )"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.Writeline "Novedad insertada "
    End If
Else 'Tiene Vigencia
    'Reviso que no se pise
    StrSql = "SELECT * FROM novemp WHERE "
    StrSql = StrSql & " concnro = " & concnro
    StrSql = StrSql & " AND tpanro = " & tpanro
    StrSql = StrSql & " AND empleado = " & Tercero
    StrSql = StrSql & " AND (nevigencia = 0 "
    StrSql = StrSql & " OR (nevigencia = -1 "
    If Not EsNulo(FechaHasta) Then
        StrSql = StrSql & " AND (nedesde <= " & ConvFecha(FechaHasta)
        StrSql = StrSql & " AND nehasta >= " & ConvFecha(FechaDesde) & ")"
        StrSql = StrSql & " OR  (nedesde <= " & ConvFecha(FechaHasta)
        StrSql = StrSql & " AND nehasta is null )))"
    Else
        StrSql = StrSql & " AND nehasta is null OR nehasta >= " & ConvFecha(FechaDesde) & "))"
    End If
    If rs_NovEmp.State = adStateOpen Then rs_NovEmp.Close
    OpenRecordset StrSql, rs_NovEmp

    If Not rs_NovEmp.EOF Then
        Flog.Writeline "No se puede insertar la novedad, las vigencias se superponen"
        InsertaError 1, 95
        Exit Sub
    Else
        If Not EsRetroactivo Then
            StrSql = "INSERT INTO novemp ("
            StrSql = StrSql & "empleado,concnro,tpanro,nevalor,nevigencia,nedesde"
            If Not EsNulo(FechaHasta) Then
                StrSql = StrSql & ",nehasta"
            End If
            StrSql = StrSql & ") VALUES (" & Tercero
            StrSql = StrSql & "," & concnro
            StrSql = StrSql & "," & tpanro
            StrSql = StrSql & "," & Monto
            StrSql = StrSql & ",-1"
            StrSql = StrSql & "," & ConvFecha(FechaDesde)
            If Not EsNulo(FechaHasta) Then
                StrSql = StrSql & "," & ConvFecha(FechaHasta)
            End If
            StrSql = StrSql & " )"
        Else
            StrSql = "INSERT INTO novemp ("
            StrSql = StrSql & "empleado,concnro,tpanro,nevalor,nevigencia,nedesde"
            If Not EsNulo(FechaHasta) Then
                StrSql = StrSql & ",nehasta"
            End If
            StrSql = StrSql & ",neretro,nepliqdesde,nepliqhasta"
            StrSql = StrSql & ") VALUES (" & Tercero
            StrSql = StrSql & "," & concnro
            StrSql = StrSql & "," & tpanro
            StrSql = StrSql & "," & Monto
            StrSql = StrSql & ",-1"
            StrSql = StrSql & "," & ConvFecha(FechaDesde)
            If Not EsNulo(FechaHasta) Then
                StrSql = StrSql & "," & ConvFecha(FechaHasta)
            End If
            StrSql = StrSql & "," & CInt(EsRetroactivo)
            StrSql = StrSql & "," & PeriodoDesde
            StrSql = StrSql & "," & PeriodoHasta
            StrSql = StrSql & " )"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.Writeline "Novedad insertada "
    End If
End If


'Cierro todo y libero
If rs_NovEmp.State = adStateOpen Then rs_NovEmp.Close
If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
If rs_TipoPar.State = adStateOpen Then rs_TipoPar.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close

Set rs_NovEmp = Nothing
Set rs_Empleado = Nothing
Set rs_Concepto = Nothing
Set rs_TipoPar = Nothing
Set rs_Periodo = Nothing
End Sub


Public Sub LineaModelo_214(ByVal strLinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta en emp_ticket y posiblemente en emptikdist (si es que hay distribucion).
' Autor      : FGZ
'              El formato es:
'                   Legajo; sigla; Monto; [catidad1; Valor1 ...[catidad5; Valor5]]
' Fecha      : '29/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim i As Integer

Dim Tercero As Long
Dim NroLegajo As Long

Dim Sigla As String
Dim Monto As Single

Dim cant(5) As Integer
Dim valor(5) As Single
Dim TikValnro(5) As Long
Dim MontoCorrecto As Boolean
Dim TickNro As Long
Dim EtikNro As Long

Dim Cantidades As Integer '0-5 Dice la cantidad de pares (cant,valor)

Dim rs_Empleado As New ADODB.Recordset
Dim rs_Ticket As New ADODB.Recordset
Dim rs_EMP_Ticket As New ADODB.Recordset
Dim rs_EMP_TikDist As New ADODB.Recordset
Dim rs_Ticket_Valor As New ADODB.Recordset
Dim rs_TikValor As New ADODB.Recordset

Cantidades = 0
' El formato es:
' Legajo; sigla; Monto; [catidad1; Valor1 ...[catidad5; Valor5]]

    'Nro de Legajo
    pos1 = 1
    pos2 = InStr(pos1, strLinea, Separador)
    If IsNumeric(Mid$(strLinea, pos1, pos2 - pos1)) Then
        NroLegajo = Mid$(strLinea, pos1, pos2 - pos1)
    Else
        InsertaError 1, 8
        Exit Sub
    End If
    
    'Sigla
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    Sigla = Mid(strLinea, pos1, pos2 - pos1)

    'Monto
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    If pos2 = 0 Then
        pos2 = Len(strLinea)
        Monto = Mid(strLinea, pos1, pos2)
    Else
        Monto = Mid(strLinea, pos1, pos2 - pos1)
               
        Cantidades = Cantidades + 1
        'Cantidad 1
        pos1 = pos2 + 1
        pos2 = InStr(pos1 + 1, strLinea, Separador)
        cant(Cantidades) = Mid(strLinea, pos1, pos2 - pos1)
        
        'Valor1
        pos1 = pos2 + 1
        pos2 = InStr(pos1 + 1, strLinea, Separador)
        If pos2 = 0 Then
           pos2 = Len(strLinea)
           valor(Cantidades) = Mid(strLinea, pos1, pos2)
        Else
           valor(Cantidades) = Mid(strLinea, pos1, pos2 - pos1)
                
            Cantidades = Cantidades + 1
           'Cantidad 2
           pos1 = pos2 + 1
           pos2 = InStr(pos1 + 1, strLinea, Separador)
           cant(Cantidades) = Mid(strLinea, pos1, pos2 - pos1)
                
           'Valor2
            pos1 = pos2 + 1
            pos2 = InStr(pos1 + 1, strLinea, Separador)
            If pos2 = 0 Then
                pos2 = Len(strLinea)
                valor(Cantidades) = Mid(strLinea, pos1, pos2)
            Else
                valor(Cantidades) = Mid(strLinea, pos1, pos2 - pos1)
                        
                Cantidades = Cantidades + 1
                'Cantidad 3
                pos1 = pos2 + 1
                pos2 = InStr(pos1 + 1, strLinea, Separador)
                cant(Cantidades) = Mid(strLinea, pos1, pos2 - pos1)
                     
                'Valor3
                 pos1 = pos2 + 1
                 pos2 = InStr(pos1 + 1, strLinea, Separador)
                 If pos2 = 0 Then
                     pos2 = Len(strLinea)
                     valor(Cantidades) = Mid(strLinea, pos1, pos2)
                 Else
                     valor(Cantidades) = Mid(strLinea, pos1, pos2 - pos1)
                     
                     Cantidades = Cantidades + 1
                    'Cantidad 4
                    pos1 = pos2 + 1
                    pos2 = InStr(pos1 + 1, strLinea, Separador)
                    cant(Cantidades) = Mid(strLinea, pos1, pos2 - pos1)
                         
                    'Valor4
                     pos1 = pos2 + 1
                     pos2 = InStr(pos1 + 1, strLinea, Separador)
                     If pos2 = 0 Then
                         pos2 = Len(strLinea)
                         valor(Cantidades) = Mid(strLinea, pos1, pos2)
                     Else
                         valor(Cantidades) = Mid(strLinea, pos1, pos2 - pos1)
                         
                         Cantidades = Cantidades + 1
                        'Cantidad 5
                        pos1 = pos2 + 1
                        pos2 = InStr(pos1 + 1, strLinea, Separador)
                        cant(Cantidades) = Mid(strLinea, pos1, pos2 - pos1)
                             
                        'Valor5
                         pos1 = pos2 + 1
                         pos2 = Len(strLinea)
                         valor(Cantidades) = Mid(strLinea, pos1, pos2)
                     End If
                 End If
            End If
        End If
    End If

' ====================================================================
'   Validar los parametros Levantados

'Que exista el legajo
StrSql = "SELECT * FROM empleado where empleg = " & NroLegajo
OpenRecordset StrSql, rs_Empleado
If rs_Empleado.EOF Then
    Flog.Writeline "No se encontro el legajo " & NroLegajo
    InsertaError 1, 8
    Exit Sub
Else
    Tercero = rs_Empleado!ternro
End If

'Con la sigla busco el TICKET.tiknro
StrSql = "SELECT * FROM ticket WHERE tiksigla = '" & Sigla & "'"
OpenRecordset StrSql, rs_Ticket
If rs_Ticket.EOF Then
    Flog.Writeline "Codigo de Ticket " & Sigla & " desconocido"
    InsertaError 2, 35
    Exit Sub
Else
    TickNro = rs_Ticket!Tiknro
End If

'Que el monto
MontoCorrecto = True
If Not IsNumeric(Monto) Then
    Flog.Writeline "El monto no es numerico " & Monto
    InsertaError 3, 5
    Exit Sub
Else
    Select Case Cantidades
    Case 0:
        MontoCorrecto = True
    Case 1:
        If Monto <> (cant(1) * valor(1)) Then
            MontoCorrecto = False
        End If
    Case 2:
        If Monto <> (cant(1) * valor(1) + cant(2) * valor(2)) Then
            MontoCorrecto = False
        End If
    Case 3:
        If Monto <> (cant(1) * valor(1) + cant(2) * valor(2) + cant(3) * valor(3)) Then
            MontoCorrecto = False
        End If
    Case 4:
        If Monto <> (cant(1) * valor(1) + cant(2) * valor(2) + cant(3) * valor(3) + cant(4) * valor(4)) Then
            MontoCorrecto = False
        End If
    Case 5:
        If Monto <> (cant(1) * valor(1) + cant(2) * valor(2) + cant(3) * valor(3) + cant(4) * valor(4) + cant(5) * valor(5)) Then
            MontoCorrecto = False
        End If
    End Select
End If
If Not MontoCorrecto Then
    Flog.Writeline "La suma de los detalles no coincide con el monto " & Monto
    InsertaError 3, 92
    Exit Sub
End If

'=============================================================
'Busco si existe el EMP_TICKET
StrSql = "SELECT * FROM emp_ticket WHERE " & _
         " tikpednro = " & TikPedNro & _
         " AND tiknro = " & TickNro & _
         " AND empleado = " & Tercero
OpenRecordset StrSql, rs_EMP_Ticket

If rs_EMP_Ticket.EOF Then
        For i = 1 To Cantidades
            'Busco el TikValnro
            StrSql = "SELECT * FROM tikvalor WHERE "
            StrSql = StrSql & " tvalmonto = " & valor(i)
            OpenRecordset StrSql, rs_TikValor
        
            If Not rs_TikValor.EOF Then
                StrSql = "SELECT * FROM ticket_valor WHERE "
                StrSql = StrSql & " tiknro = " & TickNro
                StrSql = StrSql & " AND tvalnro =" & rs_TikValor!tvalnro
                OpenRecordset StrSql, rs_Ticket_Valor
                
                If Not rs_Ticket_Valor.EOF Then
                    TikValnro(i) = rs_Ticket_Valor!TikValnro
                Else
                    Flog.Writeline "Valor no encontrado en Ticket_Valor " & rs_TikValor!tvalnro & " para el ticket " & TickNro
                    InsertaError 3 + Cantidades, 64
                    Exit Sub
                End If
            Else
                Flog.Writeline "Valor no encontrado en TIKVALOR " & valor(i)
                InsertaError 3 + Cantidades, 64
                Exit Sub
            End If
        Next i
            
        StrSql = "INSERT INTO emp_ticket ("
        StrSql = StrSql & "empleado,tiknro,tikpednro,etikfecha,etikmonto,etikcant,etikmanual"
        StrSql = StrSql & ") VALUES (" & Tercero
        StrSql = StrSql & "," & TickNro
        StrSql = StrSql & "," & TikPedNro
        StrSql = StrSql & "," & ConvFecha(Date)
        StrSql = StrSql & "," & Monto
        StrSql = StrSql & ",0,0"
        StrSql = StrSql & " )"
        objConn.Execute StrSql, , adExecuteNoRecords
    
        EtikNro = getLastIdentity(objConn, "emp_ticket")
        
        If Cantidades <> 0 Then
            'inserto los EMP_TIKDIST
            For i = 1 To Cantidades
                StrSql = "INSERT INTO emp_tikdist ("
                StrSql = StrSql & "etiknro,tikvalnro,tiknro,etikdmonto,etikdmontouni,etikdcant"
                StrSql = StrSql & ") VALUES (" & EtikNro
                StrSql = StrSql & "," & TikValnro(i)
                StrSql = StrSql & "," & TickNro
                StrSql = StrSql & "," & valor(i) * cant(i)
                StrSql = StrSql & "," & valor(i)
                StrSql = StrSql & "," & cant(i)
                StrSql = StrSql & " )"
                objConn.Execute StrSql, , adExecuteNoRecords
            Next i
        End If
        Flog.Writeline "Ticket insertado "
Else
    If Not CBool(rs_EMP_Ticket!etikmanual) And IsNull(rs_EMP_Ticket!pronro) Then
    
        For i = 1 To Cantidades
            'Busco el TikValnro
            StrSql = "SELECT * FROM tikvalor WHERE "
            StrSql = StrSql & " tvalmonto = " & valor(i)
            OpenRecordset StrSql, rs_TikValor
        
            If Not rs_TikValor.EOF Then
                StrSql = "SELECT * FROM ticket_valor WHERE "
                StrSql = StrSql & " tiknro = " & TickNro
                StrSql = StrSql & " AND tvalnro =" & rs_TikValor!tvalnro
                OpenRecordset StrSql, rs_Ticket_Valor
                
                If Not rs_Ticket_Valor.EOF Then
                    TikValnro(i) = rs_Ticket_Valor!TikValnro
                Else
                    Flog.Writeline "Valor no encontrado en Ticket_Valor " & rs_TikValor!tvalnro & " para el ticket " & TickNro
                    InsertaError 3 + Cantidades, 64
                    Exit Sub
                End If
            Else
                Flog.Writeline "Valor no encontrado en TIKVALOR " & valor(i)
                InsertaError 3 + Cantidades, 64
                Exit Sub
            End If
        Next i
    
        'Borro los EMP_TIKDIST
        StrSql = "DELETE FROM emp_tikdist WHERE etiknro =" & rs_EMP_Ticket!EtikNro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'Piso los datos del ticket
        StrSql = "UPDATE emp_ticket SET etikmonto = " & Monto
        StrSql = StrSql & " , etikfecha = " & ConvFecha(Date)
        StrSql = StrSql & " , etikcant = 0 "
        StrSql = StrSql & " WHERE tikpednro = " & TikPedNro
        StrSql = StrSql & " AND tiknro = " & TickNro
        StrSql = StrSql & " AND empleado = " & Tercero
        objConn.Execute StrSql, , adExecuteNoRecords
            
        'inserto los EMP_TIKDIST
        For i = 1 To Cantidades
            StrSql = "INSERT INTO emp_tikdist ("
            StrSql = StrSql & "etiknro,tikvalnro,tiknro,etikdmonto,etikdmontouni,etikdcant"
            StrSql = StrSql & ") VALUES (" & rs_EMP_Ticket!EtikNro
            StrSql = StrSql & "," & TikValnro(i)
            StrSql = StrSql & "," & TickNro
            StrSql = StrSql & "," & valor(i) * cant(i)
            StrSql = StrSql & "," & valor(i)
            StrSql = StrSql & "," & cant(i)
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
        Next i
        
        Flog.Writeline "Ticket Actualizado "
    Else
        Flog.Writeline "Ticket no actualizado porque es manual o fué liquidado "
    End If
End If

'cierro y libero
If rs_Ticket.State = adStateOpen Then rs_Ticket.Close
If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
If rs_EMP_Ticket.State = adStateOpen Then rs_EMP_Ticket.Close
If rs_EMP_TikDist.State = adStateOpen Then rs_EMP_TikDist.Close
If rs_Ticket_Valor.State = adStateOpen Then rs_Ticket_Valor.Close
If rs_TikValor.State = adStateOpen Then rs_TikValor.Close

Set rs_Ticket = Nothing
Set rs_Empleado = Nothing
Set rs_EMP_Ticket = Nothing
Set rs_EMP_TikDist = Nothing
Set rs_Ticket_Valor = Nothing
Set rs_TikValor = Nothing
End Sub


Public Sub LineaModelo_215(ByVal strLinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta en Acu_age.
' Autor      : FGZ
'              El formato es:
'                   Legajo; Acunro; Monto; catidad; año,mes
' Fecha      : 29/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim i As Integer

Dim Tercero As Long
Dim NroLegajo As Long
Dim acunro As Long
Dim Monto As Single
Dim Cantidad As Single
Dim Anio As Integer
Dim mes As Integer
Dim PliqNro As Long

Dim rs_Empleado As New ADODB.Recordset
Dim rs_Acumulador As New ADODB.Recordset
Dim rs_Acu_Age As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset

' El formato es:
' Legajo; Acunro; Monto; catidad; año,mes

    'Nro de Legajo
    pos1 = 1
    pos2 = InStr(pos1, strLinea, Separador)
    If IsNumeric(Mid$(strLinea, pos1, pos2 - pos1)) Then
        NroLegajo = Mid$(strLinea, pos1, pos2 - pos1)
    Else
        InsertaError 1, 8
        Exit Sub
    End If
    
    'Acumulador
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    acunro = Mid(strLinea, pos1, pos2 - pos1)

    'Monto
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    Monto = Mid(strLinea, pos1, pos2 - pos1)
               
    'Cantidad
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    Cantidad = Mid(strLinea, pos1, pos2 - pos1)
        
    'Año
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    Anio = Mid(strLinea, pos1, pos2 - pos1)
        
    'Mes
    pos1 = pos2 + 1
    pos2 = Len(strLinea)
    mes = Mid(strLinea, pos1, pos2)

' ====================================================================
'   Validar los parametros Levantados

'Que exista el legajo
StrSql = "SELECT * FROM empleado where empleg = " & NroLegajo
OpenRecordset StrSql, rs_Empleado
If rs_Empleado.EOF Then
    Flog.Writeline "No se encontro el legajo " & NroLegajo
    InsertaError 1, 8
    Exit Sub
Else
    Tercero = rs_Empleado!ternro
End If

'Valido el acumulador
StrSql = "SELECT * FROM acumulador WHERE acunro = " & acunro
OpenRecordset StrSql, rs_Acumulador
If rs_Acumulador.EOF Then
    Flog.Writeline "El Acumulador no existe " & acunro
    InsertaError 2, 52
    Exit Sub
End If

'Que el monto sea valido
If Not IsNumeric(Monto) Then
    Flog.Writeline "El monto no es numerico " & Monto
    InsertaError 3, 5
    Exit Sub
End If

'Que la cantidad sea valida
If Not IsNumeric(Cantidad) Then
    Flog.Writeline "La cantidad no es numerico " & Cantidad
    InsertaError 4, 5
    Exit Sub
End If

'Que el año sea valido
If Not IsNumeric(Anio) Then
    Flog.Writeline "El año no es numerico " & Anio
    InsertaError 5, 5
    Exit Sub
End If

'Que el mes sea valido
If Not IsNumeric(mes) Then
    Flog.Writeline "El mes no es numerico " & Anio
    InsertaError 6, 5
    Exit Sub
End If

'Busco el pliqnro correspondiente a ese año y mes
StrSql = "SELECT * FROM PERIODO WHERE pliqmes =" & mes
StrSql = StrSql & " AND pliqanio =" & Anio
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.Writeline "No existe periodo correspondiente al año " & Anio & " y mes  " & mes
    InsertaError 6, 5
    Exit Sub
Else
    PliqNro = rs_Periodo!PliqNro
End If

'=============================================================
'Busco si existe el acu_age
StrSql = "SELECT * FROM acu_age " & _
         " WHERE acunro = " & acunro & _
         " AND pliqnro = " & PliqNro & _
         " AND empage = " & Tercero
OpenRecordset StrSql, rs_Acu_Age

If rs_Acu_Age.EOF Then
        StrSql = "INSERT INTO acu_age ("
        StrSql = StrSql & "acunro,acagmonto,acagcant,empage,pliqnro"
        StrSql = StrSql & ") VALUES (" & acunro
        StrSql = StrSql & "," & Monto
        StrSql = StrSql & "," & Cantidad
        StrSql = StrSql & "," & Tercero
        StrSql = StrSql & "," & PliqNro
        StrSql = StrSql & " )"
        objConn.Execute StrSql, , adExecuteNoRecords
    
        Flog.Writeline "Acumulador insertado "
Else
        'Piso los datos del ticket
        StrSql = "UPDATE acu_age SET acagmonto = " & Monto
        StrSql = StrSql & " , acagcant = " & Cantidad
        StrSql = StrSql & " WHERE acunro = " & acunro
        StrSql = StrSql & " AND pliqnro = " & PliqNro
        StrSql = StrSql & " AND empage = " & Tercero
        objConn.Execute StrSql, , adExecuteNoRecords
            
        Flog.Writeline "Acumulador Actualizado "
End If

'cierro y libero
If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Acu_Age.State = adStateOpen Then rs_Acu_Age.Close
If rs_Acumulador.State = adStateOpen Then rs_Acumulador.Close

Set rs_Empleado = Nothing
Set rs_Periodo = Nothing
Set rs_Acu_Age = Nothing
Set rs_Acumulador = Nothing

End Sub


Public Sub LineaModelo_216(ByVal strLinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta en acu_age. Modelo para Citrusvil
'              El formato es:
'              CUIL; Acunro; Monto; catidad; año,mes
' Autor      : FGZ
' Fecha      : 29/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim i As Integer

Dim Tercero As Long
Dim CUIL As String
'Dim AcuNro As Long
Dim Monto As Single
Dim Rem_Total As Single
Dim Remu5 As Single

Dim Cantidad As Single
Dim Anio As Integer
Dim mes As Integer
Dim PliqNro As Long

Dim rs_CUIL As New ADODB.Recordset
'Dim rs_Empleado As New adodb.Recordset
Dim rs_Acumulador As New ADODB.Recordset
Dim rs_Acu_Age As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset

' El formato es:
' CUIL; Acunro; Monto; catidad; año,mes

    'CUIL
    pos1 = 1
    pos2 = 12
    If IsNumeric(Mid$(strLinea, pos1, pos2 - pos1)) Then
        CUIL = Mid$(strLinea, pos1, pos2 - pos1)
        CUIL = Left$(CUIL, 2) & "-" & Mid$(CUIL, 3, Len(CUIL) - 3) & "-" & Right$(CUIL, 1)
    Else
        InsertaError 1, 8
        Exit Sub
    End If
    
    'Monto
    'Rem total
    pos1 = 70
    pos2 = 78
    Rem_Total = CSng(Replace(Mid(strLinea, pos1, pos2 - pos1), SeparadorDecimal, "."))
               
    'Remu5
    pos1 = 298
    pos2 = 306
    Remu5 = CSng(Replace(Mid(strLinea, pos1, pos2 - pos1), SeparadorDecimal, "."))
               
    Monto = Rem_Total + Remu5
    
    'Parametros fijos ------------------------------------
    'Acumulador
'    AcuNro = 1
    'se sacó del confrep afuera del este procedimiento
    
    'Cantidad
    Cantidad = 0
    
    'Mes y año
    If Month(Date) = 1 Then
        mes = 12
        Anio = Year(Date) - 1
    Else
        mes = Month(Date) - 1
        Anio = Year(Date)
    End If
    
' ====================================================================

'   Validar los parametros Levantados
    ' Buscar el CUIL
    StrSql = " SELECT * from ter_doc cuil " & _
             " WHERE tidnro = 10 AND cuil.nrodoc = '" & CUIL & "'"
    OpenRecordset StrSql, rs_CUIL
    If Not rs_CUIL.EOF Then
        Tercero = rs_CUIL!ternro
    Else
        Flog.Writeline "No se encontró el tercero correspopndiente al cuil " & CUIL
        InsertaError 1, 8
        Exit Sub
    End If

''Valido el acumulador
'StrSql = "SELECT * FROM acumulador WHERE acunro = " & AcuNro
'OpenRecordset StrSql, rs_Acumulador
'If rs_Acumulador.EOF Then
'    Flog.writeline "El Acumulador no existe " & AcuNro
'    InsertaError 2, 35
'    Exit Sub
'End If

'Que el monto sea valido
If Not IsNumeric(Monto) Then
    Flog.Writeline "El monto no es numerico " & Monto
    InsertaError 3, 5
    Exit Sub
End If

''Que la cantidad sea valida
'If Not IsNumeric(Cantidad) Then
'    Flog.writeline "La cantidad no es numerico " & Cantidad
'    InsertaError 4, 5
'    Exit Sub
'End If
'
''Que el año sea valido
'If Not IsNumeric(Anio) Then
'    Flog.writeline "El año no es numerico " & Anio
'    InsertaError 5, 5
'    Exit Sub
'End If
'
''Que el mes sea valido
'If Not IsNumeric(Mes) Then
'    Flog.writeline "El mes no es numerico " & Anio
'    InsertaError 6, 5
'    Exit Sub
'End If

'Busco el pliqnro correspondiente a ese año y mes
StrSql = "SELECT * FROM PERIODO WHERE pliqmes =" & mes
StrSql = StrSql & " AND pliqanio =" & Anio
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.Writeline "No existe periodo correspondiente al año " & Anio & " y mes  " & mes
    InsertaError 6, 5
    Exit Sub
Else
    PliqNro = rs_Periodo!PliqNro
End If

'=============================================================
'Busco si existe el acu_age
StrSql = "SELECT * FROM acu_age " & _
         " WHERE acunro = " & acunro & _
         " AND pliqnro = " & PliqNro & _
         " AND empage = " & Tercero
OpenRecordset StrSql, rs_Acu_Age

If rs_Acu_Age.EOF Then
        StrSql = "INSERT INTO acu_age ("
        StrSql = StrSql & "acunro,acagmonto,acagcant,empage,pliqnro"
        StrSql = StrSql & ") VALUES (" & acunro
        StrSql = StrSql & "," & Monto
        StrSql = StrSql & "," & Cantidad
        StrSql = StrSql & "," & Tercero
        StrSql = StrSql & "," & PliqNro
        StrSql = StrSql & " )"
        objConn.Execute StrSql, , adExecuteNoRecords
    
        Flog.Writeline "Acumulador insertado "
Else
        'Piso los datos del ticket
        StrSql = "UPDATE acu_age SET acagmonto = " & Monto
        StrSql = StrSql & " , acagcant = " & Cantidad
        StrSql = StrSql & " WHERE acunro = " & acunro
        StrSql = StrSql & " AND pliqnro = " & PliqNro
        StrSql = StrSql & " AND empage = " & Tercero
        objConn.Execute StrSql, , adExecuteNoRecords
            
        Flog.Writeline "Acumulador Actualizado "
End If

'cierro y libero
If rs_CUIL.State = adStateOpen Then rs_CUIL.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Acu_Age.State = adStateOpen Then rs_Acu_Age.Close
If rs_Acumulador.State = adStateOpen Then rs_Acumulador.Close

Set rs_CUIL = Nothing
Set rs_Periodo = Nothing
Set rs_Acu_Age = Nothing
Set rs_Acumulador = Nothing

End Sub



Public Sub LineaModelo_217(ByVal strLinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta en vales.
'              El formato es:
'                   Legajo; Monto; Fecha ; Tipo
' Autor      : FGZ
' Fecha      : 29/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim i As Integer

Dim Tercero As Long
Dim NroLegajo As Long
Dim TipoVale As String
Dim Monto As Single
Dim FechaVale As Date
Dim DescripcionVale As String
Dim PliqNro As Long

Dim MontoCorrecto As Boolean

Dim rs_Empleado As New ADODB.Recordset
Dim rs_TipoVale As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset

' El formato es:
' Legajo; Monto; Fecha ; Tipo

    'Nro de Legajo
    pos1 = 1
    pos2 = InStr(pos1, strLinea, Separador)
    If IsNumeric(Mid$(strLinea, pos1, pos2 - pos1)) Then
        NroLegajo = Mid$(strLinea, pos1, pos2 - pos1)
    Else
        InsertaError 1, 8
        Exit Sub
    End If
    
    'Monto
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    Monto = Mid(strLinea, pos1, pos2 - pos1)
               
    'Fecha
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    FechaVale = CDate(Mid(strLinea, pos1, pos2 - pos1))
               
    pos1 = pos2 + 1
    pos2 = Len(strLinea)
    TipoVale = Mid(strLinea, pos1, pos2)

' ====================================================================
'   Validar los parametros Levantados

'Que exista el legajo
StrSql = "SELECT * FROM empleado where empleg = " & NroLegajo
OpenRecordset StrSql, rs_Empleado
If rs_Empleado.EOF Then
    Flog.Writeline "No se encontro el legajo " & NroLegajo
    InsertaError 1, 8
    Exit Sub
Else
    Tercero = rs_Empleado!ternro
End If

'Busco el periodo
StrSql = "SELECT * FROM periodo WHERE pliqmes = " & Month(FechaVale)
StrSql = StrSql & " AND pliqanio =" & Year(FechaVale)
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.Writeline "No se encontro el Periodo para el mes " & Month(FechaVale)
    InsertaError 3, 50
    Exit Sub
Else
    PliqNro = rs_Periodo!PliqNro
End If

'Que el monto
MontoCorrecto = True
If Not IsNumeric(Monto) Then
    Flog.Writeline "El monto no es numerico " & Monto
    InsertaError 2, 5
    Exit Sub
Else
    MontoCorrecto = True
End If

'Tipo de Vale
StrSql = "SELECT * FROM tipovale WHERE tvalenro = " & CLng(TipoVale)
OpenRecordset StrSql, rs_TipoVale
If rs_TipoVale.EOF Then
    Flog.Writeline "No se encontro el tipo de vale " & TipoVale
    InsertaError 2, 76
    Exit Sub
Else
    'DescripcionVale = CStr(rs_TipoVale!tvalenro) & " " & NroLegajo & " " & FechaVale
    DescripcionVale = Left(rs_TipoVale!tvaledesabr & " " & NroLegajo & " " & FechaVale, 30)
End If


'=============================================================
'Inserto el vale
StrSql = "INSERT INTO vales ("
StrSql = StrSql & "empleado,valmonto,valfecped,valfecPrev,pliqnro,pliqdto,valdesc,tvalenro,valrevis,valautoriz "
StrSql = StrSql & ") VALUES (" & Tercero
StrSql = StrSql & "," & Monto
StrSql = StrSql & "," & ConvFecha(FechaVale)
StrSql = StrSql & "," & ConvFecha(FechaVale)
StrSql = StrSql & "," & PliqNro
StrSql = StrSql & "," & PliqNro
StrSql = StrSql & ",'" & DescripcionVale & "'"
StrSql = StrSql & "," & TipoVale
StrSql = StrSql & ",0,0"
StrSql = StrSql & " )"
objConn.Execute StrSql, , adExecuteNoRecords

Flog.Writeline " Vale insertado "

'cierro y libero
If rs_TipoVale.State = adStateOpen Then rs_TipoVale.Close
If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close

Set rs_TipoVale = Nothing
Set rs_Empleado = Nothing
Set rs_Periodo = Nothing


End Sub


Public Sub LineaModelo_218(ByVal strReg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de Novedades
'              El formato es:
'                       Legajo; conccod; tpanro; monto
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer

Dim Legajo          As String   'Legajo
Dim CtoCodigo       As String   'Código de concepto
Dim Parametro       As String   'Parametro
Dim Monto           As String   'Monto

Dim NroTercero          As Integer

Dim Nro_Legajo          As Integer
Dim nro_concepto        As Integer
Dim nro_periodo         As Integer
Dim nro_proceso         As Integer
Dim nro_cabecera        As Integer
Dim nro_tipoconc        As Integer
Dim nro_tipoconc1       As Integer

Dim RsNovemp     As New ADODB.Recordset
Dim RsConcepto   As New ADODB.Recordset
Dim RsEmple   As New ADODB.Recordset

'    RegLeidos = RegLeidos + 1
    
    Flog.Writeline "Numero de Linea = " & RegLeidos
    
    ' Recupero los Valores del Archivo
    pos1 = 1
    pos2 = InStr(pos1, strReg, Separador)
    Legajo = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    CtoCodigo = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Parametro = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = Len(strReg) + 1
    Monto = Mid(strReg, pos1, pos2 - pos1)
    
' Busco al empleado asociado

  StrSql = " SELECT ternro,empleg FROM empleado WHERE empleg = " & Legajo
  OpenRecordset StrSql, RsEmple
  
  If RsEmple.EOF Then 'No existe el empleado => Error
    NroTercero = 0
  Else
    NroTercero = RsEmple!ternro
  End If
  
  If NroTercero <> 0 Then
      Flog.Writeline "Procesando al empleado = " & RsEmple!empleg
      
    ' Busco el concepto
    
      StrSql = " SELECT concnro FROM concepto WHERE conccod = '" & CtoCodigo & "'"
      OpenRecordset StrSql, RsConcepto
      
      If RsConcepto.EOF Then  'No existe el concepto => Error
        nro_concepto = 0
      Else
        nro_concepto = RsConcepto!concnro
      End If
    
      If nro_concepto <> 0 Then
      ' Busco si existe novedad de liquidacion para el mismo legajo-concepto-parametro
        StrSql = " SELECT * FROM novemp WHERE concnro = " & nro_concepto
        StrSql = StrSql & " AND empleado = " & NroTercero
        StrSql = StrSql & " AND tpanro = " & Parametro
        OpenRecordset StrSql, RsNovemp
        
        If Not RsNovemp.EOF Then  'Existe la novedad para el empleado
            StrSql = " UPDATE novemp SET nevalor = " & Monto
            StrSql = StrSql & " WHERE concnro = " & nro_concepto
            StrSql = StrSql & " AND empleado = " & NroTercero
            StrSql = StrSql & " AND tpanro = " & Parametro
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.Writeline "Actualice el Detalle de liquidacion  - " & CtoCodigo & " - " & Parametro & " - " & Monto
        Else
            StrSql = " INSERT INTO novemp(empleado,concnro,tpanro,nevalor,nevigencia) VALUES ("
            StrSql = StrSql & NroTercero & "," & nro_concepto & "," & Parametro & "," & Monto & ",0)"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.Writeline "Inserte el Detalle de liquidacion  - " & CtoCodigo & " - " & Parametro & " - " & Monto
        End If
        
      Else
        Flog.Writeline "Concepto Inexistente = " & CtoCodigo
      End If
  Else
      Flog.Writeline "Empleado Inexistente = " & RsEmple!empleg
  End If
End Sub


Public Sub LineaModelo_219(ByVal strReg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de Familiares
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Legajo          As Long ' Legajo del Empleado
Dim Apellido        As String  ' Apellido del Familiar
Dim nombre          As String  ' Nombre del Familiar
Dim Fnac            As String  ' Fecha de Nacimiento del Familiar
Dim NAC             As String  ' Nacionalidad del Familiar
Dim PaisNac         As String  ' Pais de Nacimiento
Dim EstCiv          As String  ' Estado Civil
Dim Sexo            As String  ' Sexo del Familiar
Dim GPare           As String  ' Grado de Parentesco
Dim Disc            As String  ' Discapacitado
Dim Estudia         As String  ' Estudia
Dim NivEst          As String  ' Nivel de Estudio
Dim TipDoc          As String  ' Tipo de Documento del Familiar
Dim nrodoc          As String  ' Nº de Documento del Familiar
Dim calle           As String   'Calle                    -- detdom.calle
Dim nro             As String   'Número                   -- detdom.nro
Dim piso            As String   'Piso                     -- detdom.piso
Dim Depto           As String   'Depto                    -- detdom.depto
Dim Torre           As String   'Torre                    -- detdom.torre
Dim Manzana         As String   'Manzana                  -- detdom.manzana
Dim Cpostal         As String   'Cpostal                  -- detdom.codigopostal
Dim Entre           As String   'Entre Calles             -- detdom.entrecalles
Dim Barrio          As String   'Barrio                   -- detdom.barrio
Dim Localidad       As String   'Localidad                -- detdom.locnro
Dim Partido         As String   'Partido                  -- detdom.partnro
Dim Zona            As String   'Zona                     -- detdom.zonanro
Dim Provincia       As String   'Provincia                -- detdom.provnro
Dim Pais            As String   'Pais                     -- detdom.paisnro
Dim Telefono        As String   'Telefono                 -- telefono.telnro
Dim ObraSocial      As String   'Obra Social
Dim PlanOSocial     As String   'Plan Obra Social
Dim AvisoEmer       As String   'Aviso ante Emergencia
Dim PagaSalario     As String   'Paga Salario Familiar
Dim Ganancias       As String   'Se lo toma para ganancias

Dim CUIL            As String  ' CUIL del Familiar
Dim ESC             As String  ' Escolaridad
Dim GRADO           As String  ' Grado al que concurre
Dim NroTDoc         As String

Dim pos1            As Integer
Dim pos2            As Integer

Dim NroTercero      As Integer
Dim NroEmpleado     As Integer
Dim CodTerFam       As String
Dim nro_nrodom      As Integer
Dim nro_nacionalidad As Integer
Dim nro_paisnac      As Integer
Dim nro_estciv      As Integer
Dim nro_Sexo        As Integer
Dim nro_estudia     As Integer
Dim nro_osocial     As Integer
Dim nro_planos      As Integer
Dim nro_aviso       As Integer
Dim nro_salario     As Integer
Dim nro_gan         As Integer
Dim nro_disc        As Integer
Dim nro_paren        As Integer
Dim nro_barrio          As Integer
Dim nro_localidad       As Integer
Dim nro_partido         As Integer
Dim nro_zona            As Integer
Dim nro_provincia       As Integer
Dim nro_pais            As Integer
Dim OSocial             As String
Dim ter_osocial         As Integer
Dim Inserto_estr        As Boolean


Dim StrSql          As String
Dim rs              As New ADODB.Recordset

'    RegLeidos = RegLeidos + 1
    
    Flog.Writeline "Numero de Linea = " & RegLeidos
    
    pos1 = 1
    pos2 = InStr(pos1, strReg, Separador)
    Legajo = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Apellido = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    nombre = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Fnac = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    If Fnac = "N/A" Or Fnac = " " Then
        Fnac = "''"
    Else
       Fnac = ConvFecha(Fnac)
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    PaisNac = Trim(Mid(strReg, pos1, pos2 - pos1))
    StrSql = " SELECT paisnro FROM pais WHERE paisdesc = '" & PaisNac & "'"
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        nro_paisnac = rs!paisnro
    Else
        nro_paisnac = 0
    End If
    rs.Close
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    NAC = Trim(Mid(strReg, pos1, pos2 - pos1))
    StrSql = " SELECT nacionalnro FROM nacionalidad WHERE nacionaldes = '" & NAC & "'"
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        nro_nacionalidad = rs!nacionalnro
    Else
        nro_nacionalidad = 0
    End If
    rs.Close
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    EstCiv = Trim(Mid(strReg, pos1, pos2 - pos1))
    StrSql = " SELECT estcivnro FROM estcivil WHERE estcivdesabr = '" & EstCiv & "'"
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        nro_estciv = rs!estcivnro
    Else
        nro_estciv = 0
    End If
    rs.Close
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Sexo = Trim(Mid(strReg, pos1, pos2 - pos1))
    If Sexo = "M" Or Sexo = "Masculino" Or Sexo = "MASCULINO" Then
        nro_Sexo = -1
    Else
        nro_Sexo = 0
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    GPare = Trim(Mid(strReg, pos1, pos2 - pos1))
    StrSql = " SELECT parenro FROM parentesco WHERE paredesc = '" & GPare & "'"
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        nro_paren = rs!parenro
    Else
        nro_paren = 0
    End If
    rs.Close
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Disc = Trim(Mid(strReg, pos1, pos2 - pos1))
    If Disc = "N/A" Or Disc = "NO" Then
        nro_disc = 0
    Else
        nro_disc = -1
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Estudia = Trim(Mid(strReg, pos1, pos2 - pos1))
    If Estudia = "N/A" Or Estudia = "NO" Then
        nro_estudia = 0
    Else
        nro_estudia = -1
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    NivEst = Trim(Mid(strReg, pos1, pos2 - pos1))
' Por ahora no hago nada con el nivel de estudio porque en Accor no lo pasaron

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    TipDoc = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    StrSql = " SELECT tidnro FROM tipodocu WHERE tidsigla = '" & TipDoc & "'"
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        NroTDoc = rs!tidnro
    Else
        NroTDoc = 0
    End If
    rs.Close
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    nrodoc = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    If nrodoc = "N/A" Then
        nrodoc = ""
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    calle = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    If calle = "N/A" Then
        calle = ""
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    nro = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    If nro = "N/A" Then
        nro = ""
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    piso = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    If piso = "N/A" Then
        piso = ""
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Depto = Trim(Mid(strReg, pos1, pos2 - pos1))

    If Depto = "N/A" Then
        Depto = ""
    End If

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Torre = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    If Torre = "N/A" Then
        Torre = ""
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Manzana = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    If Manzana = "N/A" Then
        Manzana = ""
    End If

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Cpostal = Trim(Mid(strReg, pos1, pos2 - pos1))

    If Cpostal = "N/A" Then
        Cpostal = ""
    End If

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Entre = Trim(Mid(strReg, pos1, pos2 - pos1))

    If Entre = "N/A" Then
        Entre = ""
    End If

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Barrio = Trim(Mid(strReg, pos1, pos2 - pos1))

    If Barrio = "N/A" Then
        Barrio = ""
    End If

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Localidad = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Partido = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Zona = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Provincia = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Pais = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    If Pais <> "N/A" Then
        Call ValidarPais(Pais, nro_pais)
    End If
    If Provincia <> "N/A" Then
        Call ValidarProvincia(Provincia, nro_provincia, nro_pais)
    End If
    If Localidad <> "N/A" Then
        Call ValidarLocalidad(Localidad, nro_localidad, nro_pais, nro_provincia)
    End If
    If Partido <> "N/A" Then
        Call ValidarPartido(Partido, nro_partido)
    End If
    If Zona <> "N/A" Then
        Call ValidarZona(Zona, nro_zona, nro_provincia)
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Telefono = Mid(strReg, pos1, pos2 - pos1)
    
    If Telefono = "N/A" Then
        Telefono = ""
    End If

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    ObraSocial = Trim(Mid(strReg, pos1, pos2 - pos1))
    If ObraSocial = "N/A" Or ObraSocial = "" Then
        nro_osocial = 0
    Else
        StrSql = " SELECT ternro FROM osocial WHERE osdesc = '" & ObraSocial & "'"
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            nro_osocial = rs!ternro
        Else
            nro_osocial = 0
        End If
        rs.Close
    End If

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    PlanOSocial = Trim(Mid(strReg, pos1, pos2 - pos1))
    If PlanOSocial = "N/A" Or PlanOSocial = "" Then
        nro_planos = 0
    Else
        If nro_osocial <> 0 Then
            StrSql = " SELECT plnro FROM planos WHERE plnom = '" & PlanOSocial & "'"
            StrSql = StrSql & " AND osocial = " & nro_osocial
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                nro_planos = rs!plnro
            Else
                nro_planos = 0
            End If
            rs.Close
        Else
            nro_planos = 0
        End If
    End If

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    AvisoEmer = Trim(Mid(strReg, pos1, pos2 - pos1))
    If AvisoEmer = "N/A" Or AvisoEmer = "NO" Then
        nro_aviso = 0
    Else
        nro_aviso = -1
    End If

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    PagaSalario = Trim(Mid(strReg, pos1, pos2 - pos1))
    If PagaSalario = "N/A" Or PagaSalario = "NO" Then
        nro_salario = 0
    Else
        nro_salario = -1
    End If

    pos1 = pos2 + 1
    pos2 = Len(strReg) + 1
    Ganancias = Trim(Mid(strReg, pos1, pos2 - pos1))
    If Ganancias = "N/A" Or Ganancias = "NO" Then
        nro_gan = 0
    Else
        nro_gan = -1
    End If

' Busco el empleado asociado

  StrSql = "SELECT ternro FROM empleado WHERE empleg = " & Legajo
  OpenRecordset StrSql, rs
  NroEmpleado = rs!ternro

  If rs.State = adStateOpen Then
    rs.Close
  End If

' Inserto el tercero asociado al familiar

  StrSql = " INSERT INTO tercero(ternom,terape,terfecnac,tersex,nacionalnro,paisnro,estcivnro)"
  StrSql = StrSql & " VALUES('" & nombre & "','" & Apellido & "'," & Fnac & "," & nro_Sexo & ","
  If nro_nacionalidad <> 0 Then
    StrSql = StrSql & nro_nacionalidad & ","
  Else
    StrSql = StrSql & "Null,"
  End If
  If nro_paisnac <> 0 Then
    StrSql = StrSql & nro_paisnac & ","
  Else
    StrSql = StrSql & "Null,"
  End If
  StrSql = StrSql & nro_estciv & ")"
  objConn.Execute StrSql, , adExecuteNoRecords

  'StrSql = " SELECT @@IDENTITY AS CodTerFam "              'SQL
  StrSql = " SELECT MAX(ternro) AS CodTerFam FROM tercero "
  OpenRecordset StrSql, rs
  NroTercero = rs!CodTerFam

  Flog.Writeline "Codigo de Tercero-Familiar = " & NroTercero

' Inserto el Familiar

  StrSql = " INSERT INTO familiar(empleado,ternro,parenro,famest,famestudia,famcernac,faminc,famsalario,famemergencia,famcargadgi,osocial,plnro,famternro)"
  StrSql = StrSql & " values(" & NroEmpleado & "," & NroTercero & "," & nro_paren & ",-1," & nro_estudia & ",0," & nro_disc & "," & nro_salario & "," & nro_aviso & "," & nro_gan & "," & nro_osocial & "," & nro_planos & ",0)"
  objConn.Execute StrSql, , adExecuteNoRecords

  Flog.Writeline "Inserte el Familiar - " & Legajo & " - " & Apellido & " - " & nombre

' Inserto el Registro correspondiente en ter_tip

  StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & NroTercero & ",3)"
  objConn.Execute StrSql, , adExecuteNoRecords

' Inserto los Documentos
  If nrodoc <> "" And nrodoc <> "N/A" And TipDoc <> "N/A" Then
    StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
    StrSql = StrSql & " VALUES(" & NroTercero & "," & NroTDoc & ",'" & nrodoc & "')"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el DU - "
  End If
  
  If rs.State = adStateOpen Then rs.Close
  
' Inserto el Domicilio
  
  If calle <> "N/A" Then
      StrSql = " INSERT INTO cabdom(tipnro,ternro,domdefault,tidonro) "
      StrSql = StrSql & " VALUES(1," & NroTercero & ",-1,2)"
      objConn.Execute StrSql, , adExecuteNoRecords
      'StrSql = " SELECT @@IDENTITY AS CodDom "                         'SQL
      StrSql = " SELECT MAX(domnro) AS CodDom FROM cabdom "              'ORACLE
      OpenRecordset StrSql, rs
      
      StrSql = " INSERT INTO detdom(domnro,calle,nro,piso,oficdepto,torre,manzana,codigopostal,"
      StrSql = StrSql & "locnro,provnro,paisnro,barrio,partnro,zonanro) "
      StrSql = StrSql & " VALUES (" & rs!CodDom & ",'" & calle & "','" & nro & "','" & piso & "','"
      StrSql = StrSql & Depto & "','" & Torre & "','" & Manzana & "','" & Cpostal & "'," & nro_localidad & ","
      StrSql = StrSql & nro_provincia & "," & nro_pais & ",'" & Barrio & "'," & nro_partido & "," & nro_zona & ")"
      objConn.Execute StrSql, , adExecuteNoRecords
      
      Flog.Writeline "Inserte el Domicilio - "
      
      If Telefono <> "" Then
        StrSql = " INSERT INTO telefono(domnro,telnro,teldefault) "
        StrSql = StrSql & " VALUES(" & rs!CodDom & ",'" & Telefono & "',-1)"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.Writeline "Inserte el Telefono - "
      End If
      
  End If
  
  If rs.State = adStateOpen Then rs.Close
End Sub

Public Sub LineaModelo_220(ByVal strReg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de Familiares 2
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Legajo          As Long ' Legajo del Empleado
Dim Apellido        As String  ' Apellido del Familiar
Dim nombre          As String  ' Nombre del Familiar
Dim NroOSL          As String
Dim NroOSE          As String
Dim OSE             As String
Dim PlanOSE         As String
Dim PlanOdon        As String
Dim Beca            As String
Dim FPC             As String
Dim Seguro          As String

Dim pos1            As Integer
Dim pos2            As Integer

Dim NroTercero      As Integer
Dim NroEmpleado     As Integer
Dim NroFamiliar     As Integer
Dim CodTerFam       As String
Dim nro_seg             As Integer
Dim Inserto_estr        As Boolean

Dim StrSql          As String
Dim rs              As New ADODB.Recordset

'    RegLeidos = RegLeidos + 1
    
    Flog.Writeline "Numero de Linea = " & RegLeidos
    
    pos1 = 1
    pos2 = InStr(pos1, strReg, Separador)
    Legajo = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Apellido = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    nombre = Trim(Mid(strReg, pos1, pos2 - pos1))

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    NroOSL = Trim(Mid(strReg, pos1, pos2 - pos1))
    If NroOSL = "N/A" Then
        NroOSL = ""
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    OSE = Trim(Mid(strReg, pos1, pos2 - pos1))
    If OSE = "N/A" Then
        OSE = ""
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    NroOSE = Trim(Mid(strReg, pos1, pos2 - pos1))
    If NroOSE = "N/A" Then
        NroOSE = ""
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    PlanOSE = Trim(Mid(strReg, pos1, pos2 - pos1))
    If PlanOSE = "N/A" Then
        PlanOSE = ""
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    PlanOdon = Trim(Mid(strReg, pos1, pos2 - pos1))
    If PlanOdon = "N/A" Then
        PlanOdon = ""
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Beca = Trim(Mid(strReg, pos1, pos2 - pos1))
    If Beca = "N/A" Then
        Beca = ""
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    FPC = Trim(Mid(strReg, pos1, pos2 - pos1))
    If FPC = "N/A" Then
        FPC = ""
    End If
    
    pos1 = pos2 + 1
    pos2 = Len(strReg) + 1
    Seguro = Trim(Mid(strReg, pos1, pos2 - pos1))
    If Seguro = "N/A" Or Seguro = "NO" Then
        nro_seg = 0
    Else
        nro_seg = -1
    End If

' Busco el empleado asociado

  StrSql = "SELECT ternro FROM empleado WHERE empleg = " & Legajo
  OpenRecordset StrSql, rs
  NroEmpleado = rs!ternro

  If rs.State = adStateOpen Then
    rs.Close
  End If
  
' Busco al familiar por el nombre y apellido

  StrSql = "SELECT familiar.ternro FROM familiar "
  StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = familiar.ternro "
  StrSql = StrSql & " WHERE familiar.empleado = " & NroEmpleado
  StrSql = StrSql & " AND tercero.terape = '" & Apellido & "'"
  StrSql = StrSql & " AND tercero.ternom = '" & nombre & "'"
  OpenRecordset StrSql, rs
  
  NroFamiliar = 0
  
  If Not rs.EOF Then
    NroFamiliar = rs!ternro
    ' Inserto las Notas
    If NroOSL <> "" Then
      StrSql = " INSERT INTO notas_ter(ternro,tnonro,notatxt)"
      StrSql = StrSql & " VALUES(" & NroFamiliar & ",26,'" & NroOSL & "')"
      objConn.Execute StrSql, , adExecuteNoRecords
      Flog.Writeline Apellido & ", " & nombre & " - Inserto NroOSL"
    End If
    If OSE <> "" Then
      StrSql = " INSERT INTO notas_ter(ternro,tnonro,notatxt)"
      StrSql = StrSql & " VALUES(" & NroFamiliar & ",27,'" & OSE & "')"
      objConn.Execute StrSql, , adExecuteNoRecords
      Flog.Writeline Apellido & ", " & nombre & " - Inserto OSE"
    End If
    If NroOSE <> "" Then
      StrSql = " INSERT INTO notas_ter(ternro,tnonro,notatxt)"
      StrSql = StrSql & " VALUES(" & NroFamiliar & ",6,'" & NroOSE & "')"
      objConn.Execute StrSql, , adExecuteNoRecords
      Flog.Writeline Apellido & ", " & nombre & " - Inserto NroOSE"
    End If
    If PlanOSE <> "" Then
      StrSql = " INSERT INTO notas_ter(ternro,tnonro,notatxt)"
      StrSql = StrSql & " VALUES(" & NroFamiliar & ",28,'" & PlanOSE & "')"
      objConn.Execute StrSql, , adExecuteNoRecords
      Flog.Writeline Apellido & ", " & nombre & " - Inserto PlanOSE"
    End If
    If PlanOdon <> "" Then
      StrSql = " INSERT INTO notas_ter(ternro,tnonro,notatxt)"
      StrSql = StrSql & " VALUES(" & NroFamiliar & ",29,'" & PlanOdon & "')"
      objConn.Execute StrSql, , adExecuteNoRecords
      Flog.Writeline Apellido & ", " & nombre & " - Inserto PlanOdon"
    End If
    If Beca <> "" Then
      StrSql = " INSERT INTO notas_ter(ternro,tnonro,notatxt)"
      StrSql = StrSql & " VALUES(" & NroFamiliar & ",30,'" & Beca & "')"
      objConn.Execute StrSql, , adExecuteNoRecords
      Flog.Writeline Apellido & ", " & nombre & " - Inserto Beca"
    End If
    If FPC <> "" Then
      StrSql = " INSERT INTO notas_ter(ternro,tnonro,notatxt)"
      StrSql = StrSql & " VALUES(" & NroFamiliar & ",31,'" & FPC & "')"
      objConn.Execute StrSql, , adExecuteNoRecords
      Flog.Writeline Apellido & ", " & nombre & " - Inserto FPC"
    End If
  End If
  
  If rs.State = adStateOpen Then
    rs.Close
  End If
  
  If NroFamiliar <> 0 Then
    ' Asigno Benef. Seguro de Vida
    StrSql = "UPDATE familiar SET fambensegvida = " & nro_seg
    StrSql = StrSql & " WHERE familiar.ternro = " & NroFamiliar
    objConn.Execute StrSql, , adExecuteNoRecords
  End If
  
  If rs.State = adStateOpen Then
    rs.Close
  End If

End Sub


Public Sub LineaModelo_221(ByVal strReg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de Empleados
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1            As Integer
Dim pos2            As Integer

Dim Legajo          As String   'LEGAJO                   -- empleado.empleg
Dim Apellido        As String   'APELLIDO                 -- empleado.terape y tercero.terape
Dim nombre          As String   'NOMBRE                   -- empleado.ternom y tercero.ternom
Dim Fnac            As String   'FNAC                     -- tercero.terfecna
Dim Nacionalidad    As String   'Nacionalidad             -- tercero.nacionalnro
Dim PaisNac         As String   'Pais de Nacimiento       -- tercero.paisnro
Dim Fing            As String   'Fec.Ingreso al Pais      -- terecro.terfecing
Dim EstCivil        As String   'Est.Civil                -- tercero.estcivnro
Dim Sexo            As String   'Sexo                     -- tercero.tersex
Dim FAlta           As String   'Fec. de Alta             -- empleado.empfaltagr y fases.altfec
Dim Estudio         As String   'Estudia                  -- empleado.empestudia
Dim NivEst          As String   'Nivel de Estudio         -- empleado.nivnro
Dim Tdocu           As String   'Tipo Documento           -- ter_dpc.tidnro (DU)
Dim Ndocu           As String   'Nro. Documento           -- ter_doc.nrodoc
Dim CUIL            As String   'CUIL                     -- ter_doc.nrodoc (10)
Dim calle           As String   'Calle                    -- detdom.calle
Dim nro             As String   'Número                   -- detdom.nro
Dim piso            As String   'Piso                     -- detdom.piso
Dim Depto           As String   'Depto                    -- detdom.depto
Dim Torre           As String   'Torre                    -- detdom.torre
Dim Manzana         As String   'Manzana                  -- detdom.manzana
Dim Cpostal         As String   'Cpostal                  -- detdom.codigopostal
Dim Entre           As String   'Entre Calles             -- detdom.entrecalles
Dim Barrio          As String   'Barrio                   -- detdom.barrio
Dim Localidad       As String   'Localidad                -- detdom.locnro
Dim Partido         As String   'Partido                  -- detdom.partnro
Dim Zona            As String   'Zona                     -- detdom.zonanro
Dim Provincia       As String   'Provincia                -- detdom.provnro
Dim Pais            As String   'Pais                     -- detdom.paisnro
Dim Telefono        As String   'Telefono                 -- telefono.telnro
Dim TelLaboral        As String   'Telefono                 -- telefono.telnro
Dim TelCelular        As String   'Telefono                 -- telefono.telnro
Dim Email           As String   'E-mail                   -- empleado.empemail
Dim Sucursal        As String   'Sucursal                 -- his_estructura.estrnro
Dim Sector          As String   'Sector                   -- his_estructura.estrnro
Dim categoria       As String   'Categoria                -- his_estructura.estrnro
Dim Puesto          As String   'Puesto                   -- his_estructura.estrnro
Dim CCosto          As String   'C.Costo                  -- his_estructura.estrnro
Dim Gerencia        As String   'Gerencia                 -- his_estructura.estrnro
Dim Departamento    As String   'Departamento             -- his_estructura.estrnro
Dim Direccion       As String   'Direccion                -- his_estructura.estrnro
Dim CajaJub         As String   'Caja de Jubilacion       -- his_estructura.estrnro
Dim Sindicato       As String   'Sindicato                -- his_estructura.estrnro
Dim OSocialLey         As String   'Obra Social              -- his_estructura.estrnro
Dim PlanOSLey          As String   'Plan OS                  -- his_estructura.estrnro
Dim OSocialElegida         As String   'Obra Social              -- his_estructura.estrnro
Dim PlanOSElegida          As String   'Plan OS                  -- his_estructura.estrnro
Dim Contrato        As String   'Contrato                 -- his_estructura.estrnro
Dim Convenio        As String   'Convenio                 -- his_estructura.estrnro
Dim LPago           As String   'Lugar de Pago            -- his_estructura.estrnro
Dim RegHorario      As String   'Regimen Horario          -- his_estructura.estrnro
Dim FormaLiq        As String   'Forma de Liquidacion     -- his_estructura.estrnro
Dim FormaPago       As String   'Forma de Pago            -- formapago.fpagdescabr
Dim SucBanco        As String   'Sucursal del Banco       -- ctabancaria.ctabsuc
Dim BancoPago       As String   'Banco Pago               -- his_estructura.estrnro, formapago.fpagbanc (siempre y cuando el Banco sea <> 0) y ctabancaria.banco
Dim NroCuenta       As String   'Nro. Cuenta              -- ctabancario.ctabnro
Dim Actividad       As String   'Actividad                -- his_estructura.estrnro
Dim CondSIJP        As String   'Condicion SIJP           -- his_estructura.estrnro
Dim SitRev          As String   'Sit. de Revista SIJP     -- his_estructura.estrnro
Dim ModCont         As String   'Mod. de Contrat. SIJP    -- his_estructura.estrnro
Dim ART             As String   'ART                      -- his_estructura.estrnro
Dim Estado          As String   'Estado                   -- empleado.empest y fases.estado
Dim CausaBaja       As String   'Causa de Baja            -- fases.caunro
Dim FBaja           As String   'Fecha de Baja            -- fases.bajfec
Dim Empresa         As String   'Empresa                  -- his_estructura.estrnro
Dim ModOrg         As String   'Empresa                  -- his_estructura.estrnro
Dim OSL        As String   'Empresa                  -- his_estructura.estrnro
Dim OSE         As String   'Empresa                  -- his_estructura.estrnro
Dim PlanOdon         As String   'Empresa                  -- his_estructura.estrnro
Dim Locacion         As String   'Empresa                  -- his_estructura.estrnro
Dim Area         As String   'Empresa                  -- his_estructura.estrnro
Dim SubDepto         As String   'Empresa                  -- his_estructura.estrnro
Dim NroCBU         As String   'Empresa                  -- his_estructura.estrnro

Dim ternro As Long

Dim NroTercero          As Integer

Dim Nro_Legajo          As Integer
Dim nro_tdocumento      As Integer
Dim nro_nivest          As Integer
Dim nro_estudio         As Integer

Dim nro_nrodom          As Integer

Dim nro_barrio          As Integer
Dim nro_localidad       As Integer
Dim nro_partido         As Integer
Dim nro_zona            As Integer
Dim nro_provincia       As Integer
Dim nro_pais            As Integer
Dim nro_paisnac         As Integer

Dim nro_sucursal        As Integer
Dim nro_sector          As Integer
Dim nro_categoria       As Integer
Dim nro_puesto          As Integer
Dim nro_ccosto          As Integer
Dim nro_gerencia        As Integer
Dim nro_cajajub         As Integer
Dim nro_sindicato       As Integer
Dim nro_osocial_ley         As Integer
Dim nro_planos_ley          As Integer
Dim nro_osocial_elegida         As Integer
Dim nro_planos_elegida          As Integer
Dim nro_contrato        As Integer
Dim nro_convenio        As Integer
Dim nro_reghorario      As Integer
Dim nro_formaliq        As Integer
Dim nro_bancopago       As Integer
Dim nro_actividad       As Integer
Dim nro_sitrev          As Integer
Dim nro_modcont         As Integer
Dim nro_art             As Integer
Dim nro_departamento    As Integer
Dim nro_direccion       As Integer
Dim nro_lpago           As Integer
Dim nro_condsijp        As Integer
Dim nro_formapago       As Integer
Dim nro_causabaja       As Integer
Dim nro_empresa         As Integer
Dim NroDom              As Integer
Dim nro_osl             As Integer
Dim nro_odon            As Integer
Dim nro_ose             As Integer
Dim nro_locacion        As Integer
Dim nro_area            As Integer
Dim nro_SubDepto        As Integer
Dim nro_ModOrg          As Integer

Dim nro_estcivil        As Integer
Dim nro_nacionalidad    As Integer

Dim F_Nacimiento        As String
Dim F_Fallecimiento     As String
Dim F_Alta              As String
Dim F_Baja              As String
Dim F_Ingreso           As String

Dim Inserto_estr        As Boolean

Dim ter_sucursal        As Integer
Dim ter_empresa         As Integer
Dim ter_cajajub         As Integer
Dim ter_sindicato       As Integer
Dim ter_osocial_ley         As Integer
Dim ter_osocial_elegida         As Integer
Dim ter_bancopago       As Integer
Dim ter_art             As Integer
Dim ter_sexo            As Integer
Dim ter_estudio         As Integer
Dim ter_estado          As Integer

Dim fpgo_bancaria       As Integer

Dim rs As New ADODB.Recordset
Dim rs_sql As New ADODB.Recordset
Dim rs_Empleado As New ADODB.Recordset
Dim rs_Tel As New ADODB.Recordset

Dim SucDesc             As Boolean   'Sucursal                 -- his_estructura
Dim SecDesc             As Boolean   'Sector                   -- his_estructura
Dim CatDesc             As Boolean   'Categoria                -- his_estructura
Dim PueDesc             As Boolean   'Puesto                   -- his_estructura
Dim CCoDesc             As Boolean   'C.Costo                  -- his_estructura
Dim GerDesc             As Boolean   'Gerencia                 -- his_estructura
Dim DepDesc             As Boolean   'Departamento             -- his_estructura
Dim DirDesc             As Boolean   'Direccion                -- his_estructura
Dim CaJDesc             As Boolean   'Caja de Jubilacion       --
Dim SinDesc             As Boolean   'Sindicato                -- his_estructura
Dim OSoElegidaDesc             As Boolean   'Obra Social              -- his_estructura
Dim PoSElegidaDesc             As Boolean   'Plan OS                  -- his_estructura
Dim OSoLeyDesc             As Boolean   'Obra Social              -- his_estructura
Dim PoSLeyDesc             As Boolean   'Plan OS                  -- his_estructura
Dim CotDesc             As Boolean   'Contrato                 -- his_estructura
Dim CovDesc             As Boolean   'Convenio                 -- his_estructura
Dim LPaDesc             As Boolean   'Lugar de Pago            -- his_estructura
Dim RegDesc             As Boolean   'Regimen Horario          -- his_estructura
Dim FLiDesc             As Boolean   'Forma de Liquidacion     -- his_estructura
Dim FPaDesc             As Boolean   'Forma de Pago            -- his_estructura
Dim BcoDesc             As Boolean   'Banco Pago               --
Dim ActDesc             As Boolean   'Actividad                --
Dim CSJDesc             As Boolean   'Condicion SIJP           --
Dim SReDesc             As Boolean   'Sit. de Revista SIJP     --
Dim MCoDesc             As Boolean   'Mod. de Contrat. SIJP    --
Dim ARTDesc             As Boolean   'ART                      --
Dim empdesc             As Boolean   'Empresa                  --
Dim OSLDesc             As Boolean   'Empresa                  --
Dim POdoDesc             As Boolean   'Empresa                  --
Dim OSEDesc             As Boolean   'Empresa                  --
Dim LocDesc             As Boolean   'Empresa                  --
Dim AreaDesc             As Boolean   'Empresa                  --
Dim SubDepDesc             As Boolean   'Empresa                  --

    ' True indica que se hace por Descripcion. False por Codigo Externo

    SucDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    SecDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    CatDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    PueDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    CCoDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    GerDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    DepDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    DirDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    CaJDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    SinDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    OSoElegidaDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    PoSElegidaDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    OSoLeyDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    PoSLeyDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    CotDesc = False ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    CovDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    LPaDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    RegDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    FLiDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    FPaDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    BcoDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    ActDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    CSJDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    SReDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    MCoDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    ARTDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    empdesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    OSLDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    POdoDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    OSEDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    LocDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    AreaDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    SubDepDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo

'    RegLeidos = RegLeidos + 1
    
    Flog.Writeline "Numero de Linea = " & RegLeidos
    
    ' Recupero los Valores del Archivo
    
    pos1 = 1
    pos2 = InStr(pos1, strReg, Separador) - 1
    Legajo = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    
    'Si el legajo ya esta me voy
    StrSql = "SELECT * FROM empleado where empleg = " & Legajo
    OpenRecordset StrSql, rs_Empleado
    If Not rs_Empleado.EOF Then
        Flog.Writeline "El Legajo ya existe " & Legajo
        InsertaError 1, 93
        Exit Sub
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Apellido = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    nombre = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Fnac = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If Fnac = "N/A" Then
       F_Nacimiento = "Null"
    Else
       F_Nacimiento = ConvFecha(Fnac)
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    PaisNac = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If PaisNac <> "N/A" Then
        StrSql = " SELECT paisnro FROM pais WHERE paisdesc = '" & PaisNac & "'"
        OpenRecordset StrSql, rs_sql
        If Not rs_sql.EOF Then
            nro_paisnac = rs_sql!paisnro
        Else
            StrSql = " INSERT INTO pais(paisdesc,paisdef) VALUES ('" & PaisNac & "',0)"
            objConn.Execute StrSql, , adExecuteNoRecords
            StrSql = " SELECT MAX(paisnro) AS MaxSql FROM pais "      ' Oracle
            'StrSql = " SELECT @@IDENTITY AS MaxSql "                    ' SQL
            OpenRecordset StrSql, rs_sql
            nro_paisnac = rs_sql!MaxSql
        End If
    Else
        nro_paisnac = 0
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Nacionalidad = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If Nacionalidad <> "N/A" Then
        StrSql = " SELECT nacionalnro FROM nacionalidad WHERE nacionaldes = '" & Nacionalidad & "'"
        OpenRecordset StrSql, rs_sql
        If Not rs_sql.EOF Then
            nro_nacionalidad = rs_sql!nacionalnro
        Else
            StrSql = " INSERT INTO nacionalidad(nacionaldes) VALUES ('" & Nacionalidad & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            StrSql = " SELECT MAX(nacionalnro) AS MaxSql FROM nacionalidad "      ' Oracle
            'StrSql = " SELECT @@IDENTITY AS MaxSql "                    ' SQL
            OpenRecordset StrSql, rs_sql
            nro_nacionalidad = rs_sql!MaxSql
        End If
    Else
        nro_nacionalidad = 0
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Fing = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If (Fing = "N/A") Then
        F_Ingreso = "Null"
    Else
        F_Ingreso = ConvFecha(Fing)
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    EstCivil = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If EstCivil <> "N/A" Then
        StrSql = " SELECT estcivnro FROM estcivil WHERE estcivdesabr = '" & EstCivil & "'"
        OpenRecordset StrSql, rs_sql
        If Not rs_sql.EOF Then
            nro_estcivil = rs_sql!estcivnro
        Else
            StrSql = " INSERT INTO estcivil(estcivdesabr) VALUES ('" & EstCivil & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            StrSql = " SELECT MAX(estcivnro) AS MaxSql FROM estcivil "      ' Oracle
            'StrSql = " SELECT @@IDENTITY AS MaxSql "                    ' SQL
            OpenRecordset StrSql, rs_sql
            nro_estcivil = rs_sql!MaxSql
        End If
    Else
        nro_estcivil = 0
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Sexo = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If (Sexo = "M") Or (Sexo = "Masculino") Or (Sexo = "-1") Or (Sexo = "MASCULINO") Then
        ter_sexo = -1
    Else
        ter_sexo = 0
    End If
                                                            
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    FAlta = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If FAlta = "N/A" Then
        F_Alta = "Null"
    Else
        F_Alta = ConvFecha(FAlta)
    End If
   
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Estudio = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If Estudio <> "N/A" Then
        If Estudio = "SI" Then
            ter_estudio = -1
        Else
            ter_estudio = 0
        End If
    Else
        ter_estudio = 0
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    NivEst = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If NivEst <> "N/A" Then
        StrSql = " SELECT nivnro FROM nivest WHERE nivdesc = '" & NivEst & "'"
        OpenRecordset StrSql, rs_sql
        If Not rs_sql.EOF Then
            nro_nivest = rs_sql!Nivnro
        Else
            StrSql = " INSERT INTO nivest(nivdesc,nivsist,nivobligatorio,nivestfli) VALUES ('" & NivEst & "',-1,0,-1)"
            objConn.Execute StrSql, , adExecuteNoRecords
            StrSql = " SELECT MAX(nivnro) AS MaxSql FROM nivest "      ' Oracle
            'StrSql = " SELECT @@IDENTITY AS MaxSql "                    ' SQL
            OpenRecordset StrSql, rs_sql
            nro_nivest = rs_sql!MaxSql
        End If
    Else
        nro_nivest = 0
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Tdocu = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    
    If Tdocu <> "N/A" Then
        StrSql = " SELECT tidnro FROM tipodocu WHERE tidsigla = '" & Tdocu & "'"
        OpenRecordset StrSql, rs_sql
        If Not rs_sql.EOF Then
            nro_tdocumento = rs_sql!tidnro
        Else
            StrSql = " INSERT INTO tipodocu(tidnom,tidsigla,tidsist,instnro,tidunico) VALUES ('" & Tdocu & "','" & Tdocu & "',0,0,0)"
            objConn.Execute StrSql, , adExecuteNoRecords
            StrSql = " SELECT MAX(tidnro) AS MaxSql FROM tipodocu "      ' Oracle
            'StrSql = " SELECT @@IDENTITY AS MaxSql "                    ' SQL
            OpenRecordset StrSql, rs_sql
            nro_tdocumento = rs_sql!MaxSql
        End If
    Else
        nro_tdocumento = 0
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Ndocu = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If Ndocu = "N/A" Then
        Ndocu = ""
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    CUIL = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If CUIL = "N/A" Then
        CUIL = ""
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    calle = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If calle = "N/A" Then
        calle = ""
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    nro = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If (nro <> "N/A" And nro <> "S/N") Then
        nro_nrodom = CLng(nro)
    Else
        nro_nrodom = 0
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    piso = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If piso = "N/A" Then
        piso = ""
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Depto = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Depto = "N/A" Then
        Depto = ""
    End If

    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Torre = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If Torre = "N/A" Then
        Torre = ""
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Manzana = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If Manzana = "N/A" Then
        Manzana = ""
    End If

    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Cpostal = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Cpostal = "N/A" Then
        Cpostal = ""
    End If

    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Entre = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Entre = "N/A" Then
        Entre = ""
    End If

    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Barrio = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Barrio = "N/A" Then
        Barrio = ""
    End If

    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Localidad = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Partido = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Zona = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Provincia = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Pais = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If Pais <> "N/A" Then
        Call ValidarPais(Pais, nro_pais)
    Else
        nro_pais = 0
    End If
    
    If Provincia <> "N/A" Then
        Call ValidarProvincia(Provincia, nro_provincia, nro_pais)
    Else
        nro_provincia = 0
    End If
    
    If Localidad <> "N/A" Then
        Call ValidarLocalidad(Localidad, nro_localidad, nro_pais, nro_provincia)
    Else
        nro_localidad = 0
    End If
    
    If Partido <> "N/A" Then
        Call ValidarPartido(Partido, nro_partido)
    Else
        nro_partido = 0
    End If
    
    If Zona <> "N/A" Then
        Call ValidarZona(Zona, nro_zona, nro_provincia)
    Else
        nro_zona = 0
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Telefono = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If Telefono = "N/A" Then
        Telefono = ""
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    TelLaboral = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If TelLaboral = "N/A" Then
        TelLaboral = ""
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    TelCelular = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If TelCelular = "N/A" Then
        TelCelular = ""
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Email = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Email = "N/A" Then
        Email = ""
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Sucursal = Mid(strReg, pos1, pos2 - pos1 + 1)

    ' Validacion y Creacion de la Sucursal (junto con sus Complementos)

    If Sucursal <> "N/A" Then
        If SucDesc Then
            Call ValidaEstructura(1, Sucursal, nro_sucursal, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(1, Sucursal, nro_sucursal, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaTercero(10, Sucursal, ter_sucursal)
            Call CreaComplemento(1, ter_sucursal, nro_sucursal, Sucursal)
            Inserto_estr = False
        End If
    Else
        nro_sucursal = 0
    End If
    
    ' Validacion y Creacion del Sector
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Sector = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Sector <> "N/A" Then
        If SecDesc Then
            Call ValidaEstructura(2, Sector, nro_sector, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(2, Sector, nro_sector, Inserto_estr)
        End If
    Else
        nro_sector = 0
    End If

    ' Validacion, Creacion del Convenio
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Convenio = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Convenio <> "N/A" Then
        If CovDesc Then
            Call ValidaEstructura(19, Convenio, nro_convenio, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(19, Convenio, nro_convenio, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaComplemento(19, 0, nro_convenio, Convenio)
            Inserto_estr = False
        End If
    Else
        nro_convenio = 0
    End If
    
    ' Validacion y Creacion de la Categoria

    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    categoria = Mid(strReg, pos1, pos2 - pos1 + 1)

    If (categoria <> "N/A" And nro_convenio <> 0) Then
        If CatDesc Then
            Call ValidaEstructura(3, categoria, nro_categoria, Inserto_estr)
            'Call ValidaCategoria(3, categoria, nro_convenio, nro_categoria, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(3, categoria, nro_categoria, Inserto_estr)
        End If
    Else
        nro_categoria = 0
    End If
    
    ' Validacion y Creacion del Puesto (junto con sus Complementos)

    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Puesto = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Puesto <> "N/A" Then
        If PueDesc Then
            Call ValidaEstructura(4, Puesto, nro_puesto, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(4, Puesto, nro_puesto, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaComplemento(4, 0, nro_puesto, Puesto)
            Inserto_estr = False
        End If
    Else
        nro_puesto = 0
    End If

    ' Validacion y Creacion del Centro de Costo
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    CCosto = Mid(strReg, pos1, pos2 - pos1 + 1)

    If CCosto <> "N/A" Then
        If CCoDesc Then
            Call ValidaEstructura(5, CCosto, nro_ccosto, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(5, CCosto, nro_ccosto, Inserto_estr)
        End If
    Else
        nro_ccosto = 0
    End If

    ' Validacion y Creacion de la Gerencia
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Gerencia = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Gerencia <> "N/A" Then
        If GerDesc Then
            Call ValidaEstructura(6, Gerencia, nro_gerencia, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(6, Gerencia, nro_gerencia, Inserto_estr)
        End If
    Else
        nro_gerencia = 0
    End If

    ' Validacion y Creacion del Departamento
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Departamento = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Departamento <> "N/A" Then
        If DepDesc Then
            Call ValidaEstructura(9, Departamento, nro_departamento, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(9, Departamento, nro_departamento, Inserto_estr)
        End If
    Else
        nro_departamento = 0
    End If

    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Direccion = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Direccion <> "N/A" Then
        If DirDesc Then
            Call ValidaEstructura(35, Direccion, nro_direccion, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(35, Direccion, nro_direccion, Inserto_estr)
        End If
    Else
        nro_direccion = 0
    End If
    
    ' Validacion y Creacion de la Caja de Jubilacion
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    CajaJub = Mid(strReg, pos1, pos2 - pos1 + 1)

    If CajaJub <> "N/A" Then
        If CaJDesc Then
            Call ValidaEstructura(15, CajaJub, nro_cajajub, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(15, CajaJub, nro_cajajub, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaTercero(6, CajaJub, ter_cajajub)
            Call CreaComplemento(15, ter_cajajub, nro_cajajub, CajaJub)
        End If
    Else
        nro_cajajub = 0
    End If

    ' Validacion y Creacion del Sindicato
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Sindicato = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Sindicato <> "N/A" Then
        If SinDesc Then
            Call ValidaEstructura(16, Sindicato, nro_sindicato, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(16, Sindicato, nro_sindicato, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaTercero(5, Sindicato, ter_sindicato)
            Call CreaComplemento(16, ter_sindicato, nro_sindicato, Sindicato)
        End If
    Else
        nro_sindicato = 0
    End If
    
    ' Validacion y Creacion de la Obra Social por Ley
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    OSocialLey = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))

    If OSocialLey <> "N/A" Then
        If OSoLeyDesc Then
            Call ValidaEstructura(24, OSocialLey, nro_osocial_ley, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(24, OSocialLey, nro_osocial_ley, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaTercero(4, OSocialLey, ter_osocial_ley)
            Call CreaComplemento(24, ter_osocial_ley, nro_osocial_ley, OSocialLey)
        Else
            StrSql = " SELECT origen FROM replica_estr WHERE estrnro = " & nro_osocial_ley
            OpenRecordset StrSql, rs_sql
            ter_osocial_ley = rs_sql!Origen
            rs_sql.Close
        End If
    Else
        nro_osocial_ley = 0
    End If

    ' Validacion y Creacion del Plan de Obra Social por Ley
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    PlanOSLey = Mid(strReg, pos1, pos2 - pos1 + 1)

    If (PlanOSLey <> "N/A" And nro_osocial_ley <> 0) Then
        If PoSLeyDesc Then
            Call ValidaEstructura(25, PlanOSLey, nro_planos_ley, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(25, PlanOSLey, nro_planos_ley, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaComplemento(25, ter_osocial_ley, nro_planos_ley, PlanOSLey)
            Inserto_estr = False
        End If
    Else
        nro_planos_ley = 0
    End If
    
    ' Validacion y Creacion de la Obra Social Elegida
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    OSocialElegida = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))

    If OSocialElegida <> "N/A" Then
        If OSoElegidaDesc Then
            Call ValidaEstructura(17, OSocialElegida, nro_osocial_elegida, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(17, OSocialElegida, nro_osocial_elegida, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaTercero(4, OSocialElegida, ter_osocial_elegida)
            Call CreaComplemento(17, ter_osocial_elegida, nro_osocial_elegida, OSocialElegida)
        Else
            StrSql = " SELECT origen FROM replica_estr WHERE estrnro = " & nro_osocial_elegida
            OpenRecordset StrSql, rs_sql
            ter_osocial_elegida = rs_sql!Origen
            rs_sql.Close
        End If
    Else
        nro_osocial_elegida = 0
    End If

    ' Validacion y Creacion del Plan de Obra Social Elegida
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    PlanOSElegida = Mid(strReg, pos1, pos2 - pos1 + 1)

    If (PlanOSElegida <> "N/A" And nro_osocial_elegida <> 0) Then
        If PoSElegidaDesc Then
            Call ValidaEstructura(23, PlanOSElegida, nro_planos_elegida, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(23, PlanOSElegida, nro_planos_elegida, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaComplemento(23, ter_osocial_elegida, nro_planos_elegida, PlanOSElegida)
            Inserto_estr = False
        End If
    Else
        nro_planos_elegida = 0
    End If
    
    ' Validacion y Creacion del Contrato

    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Contrato = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Contrato <> "N/A" Then
        If CotDesc Then
            Call ValidaEstructura(18, Contrato, nro_contrato, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(18, Contrato, nro_contrato, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaComplemento(18, 0, nro_contrato, Contrato)
            Inserto_estr = False
        End If
    Else
        nro_contrato = 0
    End If
    
    ' Validacion y Creacion del Lugar de Pago
        
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    LPago = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))

    If LPago <> "N/A" Then
        If LPaDesc Then
            Call ValidaEstructura(20, LPago, nro_lpago, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(20, LPago, nro_lpago, Inserto_estr)
        End If
    Else
        nro_lpago = 0
    End If

    ' Validacion y Creacion del Regimen Horario
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    RegHorario = Mid(strReg, pos1, pos2 - pos1 + 1)

    If RegHorario <> "N/A" Then
        If RegDesc Then
            Call ValidaEstructura(21, RegHorario, nro_reghorario, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(21, RegHorario, nro_reghorario, Inserto_estr)
        End If
    Else
        nro_reghorario = 0
    End If

    ' Validacion y Creacion de la Forma de Liquidacion
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    FormaLiq = Mid(strReg, pos1, pos2 - pos1 + 1)

    If FormaLiq <> "N/A" Then
        If FLiDesc Then
            Call ValidaEstructura(22, FormaLiq, nro_formaliq, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(22, FormaLiq, nro_formaliq, Inserto_estr)
        End If
    Else
        nro_formaliq = 0
    End If

    ' Validacion y Creacion de la Forma de Pago
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    FormaPago = Mid(strReg, pos1, pos2 - pos1 + 1)

    If FormaPago <> "N/A" Then
        StrSql = " SELECT fpagnro FROM formapago WHERE fpagdescabr = '" & FormaPago & "'"
        OpenRecordset StrSql, rs_sql
        If Not rs_sql.EOF Then
            nro_formapago = rs_sql!fpagnro
        Else
            StrSql = " INSERT INTO formapago(fpagdescabr,fpagbanc,acunro,monnro) VALUES ('" & FormaPago & "',0,6,1)"
            objConn.Execute StrSql, , adExecuteNoRecords
            StrSql = " SELECT MAX(fpagnro) AS MaxSql FROM formapago "      ' Oracle
            'StrSql = " SELECT @@IDENTITY AS MaxSql "                    ' SQL
            OpenRecordset StrSql, rs_sql
            nro_formapago = rs_sql!MaxSql
        End If
    Else
        nro_formapago = 0
    End If
    
    ' Validacion y Creacion de los Bancos de Pago
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    BancoPago = Mid(strReg, pos1, pos2 - pos1 + 1)

    If BancoPago <> "N/A" Then
        If BcoDesc Then
            Call ValidaEstructura(41, BancoPago, nro_bancopago, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(41, BancoPago, nro_bancopago, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaTercero(13, BancoPago, ter_bancopago)
            Call CreaComplemento(41, ter_bancopago, nro_bancopago, BancoPago)
        End If
        fpgo_bancaria = -1
    Else
        nro_bancopago = 0
        fpgo_bancaria = 0
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    NroCuenta = Mid(strReg, pos1, pos2 - pos1 + 1)
    If NroCuenta = "N/A" Then
        NroCuenta = ""
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    NroCBU = Mid(strReg, pos1, pos2 - pos1 + 1)
    If NroCBU = "N/A" Then
        NroCBU = ""
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    SucBanco = Mid(strReg, pos1, pos2 - pos1 + 1)
    If SucBanco = "N/A" Then
        SucBanco = ""
    End If

    ' Validacion y Creacion de la Actividad
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Actividad = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Actividad <> "N/A" Then
        If ActDesc Then
            Call ValidaEstructura(29, Actividad, nro_actividad, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(29, Actividad, nro_actividad, Inserto_estr)
        End If
    Else
        nro_actividad = 0
    End If

    ' Validacion y Creacion de la Condicion SIJP
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    CondSIJP = Mid(strReg, pos1, pos2 - pos1 + 1)

    If CondSIJP <> "N/A" Then
        If CSJDesc Then
            Call ValidaEstructura(31, CondSIJP, nro_condsijp, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(31, CondSIJP, nro_condsijp, Inserto_estr)
        End If
    Else
        nro_condsijp = 0
    End If

    ' Validacion y Creacion de la Situacion de Revista
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    SitRev = Mid(strReg, pos1, pos2 - pos1 + 1)

    If SitRev <> "N/A" Then
        If SReDesc Then
            Call ValidaEstructura(30, SitRev, nro_sitrev, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(30, SitRev, nro_sitrev, Inserto_estr)
        End If
    Else
        nro_sitrev = 0
    End If
    
    ' Validacion y Creacion de la ART
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    ART = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If ART <> "N/A" Then
        If ARTDesc Then
            Call ValidaEstructura(40, ART, nro_art, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(40, ART, nro_art, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaTercero(8, ART, ter_art)
            Call CreaComplemento(40, ter_art, nro_art, ART)
        End If
    Else
        nro_art = 0
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Estado = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    
    If UCase(Estado) = "ACTIVO" Then
        ter_estado = -1
    Else
        ter_estado = 0
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    CausaBaja = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If Not EsNulo(CausaBaja) And CausaBaja <> "N/A" Then
        StrSql = " SELECT caunro FROM causa WHERE caudes = '" & CausaBaja & "'"
        OpenRecordset StrSql, rs_sql
        If Not rs_sql.EOF Then
            nro_causabaja = rs_sql!caunro
        Else
            StrSql = " INSERT INTO causa(caudes,causist,caudesvin,empnro) VALUES ('" & CausaBaja & "',0,-1,1)"
            objConn.Execute StrSql, , adExecuteNoRecords
            StrSql = " SELECT MAX(caunro) AS MaxSql FROM causa "      ' Oracle
            'StrSql = " SELECT @@IDENTITY AS MaxSql "                    ' SQL
            OpenRecordset StrSql, rs_sql
            nro_causabaja = rs_sql!MaxSql
        End If
    Else
        nro_causabaja = 0
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    FBaja = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If EsNulo(FBaja) Or FBaja = "N/A" Then
        F_Baja = "Null"
    Else
        F_Baja = ConvFecha(FBaja)
    End If
    
    ' Validacion y Creacion de la Empresa
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    If pos2 > 0 Then
        Empresa = Mid(strReg, pos1, pos2 - pos1 + 1)
    Else
        pos2 = Len(strReg)
        Empresa = Mid(strReg, pos1, pos2 - pos1 + 1)
    End If

    If Empresa <> "N/A" Then
        If Not empdesc Then
            Call ValidaEstructura(10, Empresa, nro_empresa, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(10, Empresa, nro_empresa, Inserto_estr)
        End If
    Else
        nro_empresa = 0
    End If
    
    ' Modelo de Organizacion
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    
    If pos2 > 0 Then
        ModOrg = Mid(strReg, pos1, pos2 - pos1 + 1)
        
        ' Validacion y Creacion Socio OSL
        pos1 = pos2 + 2
        pos2 = InStr(pos1, strReg, Separador) - 1
        
        If pos2 > 0 Then
            OSL = Mid(strReg, pos1, pos2 - pos1 + 1)
        
            If OSL <> "N/A" Then
                If OSLDesc Then
                    Call ValidaEstructura(49, OSL, nro_osl, Inserto_estr)
                Else
                    Call ValidaEstructuraCodExt(49, OSL, nro_osl, Inserto_estr)
                End If
            Else
                nro_osl = 0
            End If
            
            ' Validacion y Creacion del Plan Odontologico
            pos1 = pos2 + 2
            pos2 = InStr(pos1, strReg, Separador) - 1
            If pos2 > 0 Then
                PlanOdon = Mid(strReg, pos1, pos2 - pos1 + 1)
            
                If PlanOdon <> "N/A" Then
                    If POdoDesc Then
                        Call ValidaEstructura(47, PlanOdon, nro_odon, Inserto_estr)
                    Else
                        Call ValidaEstructuraCodExt(47, PlanOdon, nro_odon, Inserto_estr)
                    End If
                Else
                    nro_odon = 0
                End If
                
                ' Validacion y Creacion Socio OSE
                pos1 = pos2 + 2
                pos2 = InStr(pos1, strReg, Separador) - 1
                If pos2 > 0 Then
                    OSE = Mid(strReg, pos1, pos2 - pos1 + 1)
                
                    If OSE <> "N/A" Then
                        If OSEDesc Then
                            Call ValidaEstructura(48, OSE, nro_ose, Inserto_estr)
                        Else
                            Call ValidaEstructuraCodExt(48, OSE, nro_ose, Inserto_estr)
                        End If
                    Else
                        nro_ose = 0
                    End If
                    
                    ' Validacion y Creacion de la Locacion
                    pos1 = pos2 + 2
                    pos2 = InStr(pos1, strReg, Separador) - 1
                    If pos2 > 0 Then
                        Locacion = Mid(strReg, pos1, pos2 - pos1 + 1)
                    
                        If Locacion <> "N/A" Then
                            If LocDesc Then
                                Call ValidaEstructura(44, Locacion, nro_locacion, Inserto_estr)
                            Else
                                Call ValidaEstructuraCodExt(44, Locacion, nro_locacion, Inserto_estr)
                            End If
                        Else
                            nro_locacion = 0
                        End If
                        
                        ' Validacion y Creacion del Area
                        pos1 = pos2 + 2
                        pos2 = InStr(pos1, strReg, Separador) - 1
                        If pos2 > 0 Then
                            Area = Mid(strReg, pos1, pos2 - pos1 + 1)
                        
                            If Area <> "N/A" Then
                                If AreaDesc Then
                                    Call ValidaEstructura(45, Area, nro_area, Inserto_estr)
                                Else
                                    Call ValidaEstructuraCodExt(45, Area, nro_area, Inserto_estr)
                                End If
                            Else
                                nro_area = 0
                            End If
                            
                            ' Validacion y Creacion del SubDepartamento
                            pos1 = pos2 + 2
                            pos2 = Len(strReg) + 1
                            If pos2 > 0 Then
                                SubDepto = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
                            
                                If SubDepto <> "N/A" Then
                                    If SubDepDesc Then
                                        Call ValidaEstructura(46, SubDepto, nro_SubDepto, Inserto_estr)
                                    Else
                                        Call ValidaEstructuraCodExt(46, SubDepto, nro_SubDepto, Inserto_estr)
                                    End If
                                Else
                                    nro_SubDepto = 0
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        nro_ModOrg = 0
    End If


' Inserto el Tercero
  If F_Nacimiento = "Null" Then
    F_Nacimiento = "''"
  End If
  If F_Ingreso = "Null" Then
    F_Ingreso = "''"
  End If

  StrSql = " INSERT INTO tercero(ternom,terape,terfecnac,tersex,terestciv,estcivnro,nacionalnro,paisnro,terfecing)"
  StrSql = StrSql & " VALUES('" & nombre & "','" & Apellido & "'," & F_Nacimiento & "," & ter_sexo & "," & nro_estcivil & "," & nro_estcivil & ","
  If nro_nacionalidad <> 0 Then
    StrSql = StrSql & nro_nacionalidad & ","
  Else
    StrSql = StrSql & "null,"
  End If
  If nro_paisnac <> 0 Then
    StrSql = StrSql & nro_paisnac & ","
  Else
    StrSql = StrSql & "null,"
  End If
  StrSql = StrSql & F_Ingreso & ")"
  objConn.Execute StrSql, , adExecuteNoRecords

'  StrSql = " SELECT MAX(ternro) AS MaxSql FROM tercero "      ' Oracle
'  'StrSql = " SELECT @@IDENTITY AS MaxSql "                   ' SQL
'
'  OpenRecordset StrSql, rs
'
'  NroTercero = rs!MaxSql
  NroTercero = getLastIdentity(objConn, "tercero")

  Flog.Writeline "Codigo de Tercero = " & NroTercero

' Inserto el Empleado
  'If F_Alta = "Null" Then
  '  F_Alta = "''"
  'End If
  'If F_Baja = "Null" Then
  '  F_Baja = "''"
  'End If

   StrSql = " INSERT INTO empleado(empleg,empfecalta,empfecbaja,empest,empfaltagr,"
   StrSql = StrSql & "ternro,nivnro,empestudia,terape,ternom,empinterno,empemail,"
   StrSql = StrSql & "empnro,tplatenro) VALUES("
   StrSql = StrSql & Legajo & "," & F_Alta & "," & F_Baja & "," & ter_estado & "," & F_Alta & ","
   StrSql = StrSql & NroTercero & "," & nro_nivest & "," & ter_estudio & ",'" & Apellido & "','"
   StrSql = StrSql & nombre & "',Null,'" & Email & "',1," & nro_ModOrg & ")"
   objConn.Execute StrSql, , adExecuteNoRecords

  Flog.Writeline "Inserte el Empleado - " & Legajo & " - " & Apellido & " - " & nombre

' Inserto el Registro correspondiente en ter_tip

  StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & NroTercero & ",1)"
  objConn.Execute StrSql, , adExecuteNoRecords

' Inserto los Documentos
    
  If nro_tdocumento <> 0 Then
      StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
      StrSql = StrSql & " VALUES(" & NroTercero & "," & nro_tdocumento & ",'" & Ndocu & "')"
      objConn.Execute StrSql, , adExecuteNoRecords
      Flog.Writeline "Inserte el DU - "
  End If
  

  If CUIL <> "" Then
    StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
    StrSql = StrSql & " VALUES(" & NroTercero & ",10,'" & CUIL & "')"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el CUIL - "
  End If

' Inserto el Domicilio

  If rs.State = adStateOpen Then
    rs.Close
  End If
  
  If (nro_localidad <> 0 And nro_provincia <> 0 And nro_pais <> 0) Then
      StrSql = " INSERT INTO cabdom(tipnro,ternro,domdefault,tidonro) "
      StrSql = StrSql & " VALUES(1," & NroTercero & ",-1,2)"
      objConn.Execute StrSql, , adExecuteNoRecords
      
'      StrSql = " SELECT MAX(domnro) AS MaxSql FROM cabdom "      ' Oracle
'      'StrSql = " SELECT @@IDENTITY AS MaxSql "                    ' SQL
'      OpenRecordset StrSql, rs
'
'      NroDom = rs!MaxSql
      NroDom = getLastIdentity(objConn, "cabdom")
    
      StrSql = " INSERT INTO detdom(domnro,calle,nro,piso,oficdepto,torre,manzana,codigopostal,entrecalles,"
      StrSql = StrSql & "locnro,provnro,paisnro,barrio,partnro,zonanro) "
      StrSql = StrSql & " VALUES (" & NroDom & ",'" & calle & "'," & nro_nrodom & ",'" & piso & "','"
      StrSql = StrSql & Depto & "','" & Torre & "','" & Manzana & "','" & Cpostal & "','" & Entre & "'," & nro_localidad & ","
      StrSql = StrSql & nro_provincia & "," & nro_pais & ",'" & Barrio & "'," & nro_partido & "," & nro_zona & ")"
      objConn.Execute StrSql, , adExecuteNoRecords
    
      Flog.Writeline "Inserte el Domicilio - "
      
      If Telefono <> "" Then
        StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular) "
        StrSql = StrSql & " VALUES(" & NroDom & ",'" & Telefono & "',0,-1,0)"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.Writeline "Inserte el Telefono - "
      End If
      If TelLaboral <> "" Then
        StrSql = "SELECT * FROM telefono "
        StrSql = StrSql & " WHERE domnro =" & NroDom
        StrSql = StrSql & " AND telnro ='" & TelLaboral & "'"
        If rs_Tel.State = adStateOpen Then rs_Tel.Close
        OpenRecordset StrSql, rs_Tel
        If rs_Tel.EOF Then
            StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular) "
            StrSql = StrSql & " VALUES(" & NroDom & ",'" & TelLaboral & "',0,0,0)"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.Writeline "Inserte el Telefono Laboral - "
        End If
      End If
      If TelCelular <> "" Then
            StrSql = "SELECT * FROM telefono "
            StrSql = StrSql & " WHERE domnro =" & NroDom
            StrSql = StrSql & " AND telnro ='" & TelCelular & "'"
            If rs_Tel.State = adStateOpen Then rs_Tel.Close
            OpenRecordset StrSql, rs_Tel
            If rs_Tel.EOF Then
                StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular) "
                StrSql = StrSql & " VALUES(" & NroDom & ",'" & TelCelular & "',0,0,-1)"
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.Writeline "Inserte el Telefono Celular - "
            End If
      End If
  End If
  
  ' Inserto las Fases
  StrSql = " INSERT INTO fases(empleado,altfec,bajfec,caunro,estado,sueldo,vacaciones,indemnizacion,real,fasrecofec)"
  StrSql = StrSql & " VALUES( " & NroTercero & "," & F_Alta & "," & F_Baja & ","
  If nro_causabaja <> 0 Then
    StrSql = StrSql & nro_causabaja
  Else
    StrSql = StrSql & "null"
  End If
  StrSql = StrSql & "," & ter_estado & ",-1,-1,-1,-1,-1)"
  objConn.Execute StrSql, , adExecuteNoRecords
  
  'Inserto la cuenta bancaria
  If (nro_formapago <> 0 And nro_bancopago <> 0 And NroCuenta <> "") Then
    StrSql = " INSERT INTO ctabancaria (ternro,fpagnro,banco,ctabestado,"
    StrSql = StrSql & "ctabsuc,ctabnro,ctabporc,ctabcbu) VALUES ("
    StrSql = StrSql & NroTercero & "," & nro_formapago & "," & nro_bancopago & ","
    StrSql = StrSql & "-1,'" & SucBanco & "','" & NroCuenta & "',100,'" & NroCBU & "')"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte la Cuenta Bancaria - "
  End If
             
  ' Inserto las Estructuras
  
  Call AsignarEstructura(1, nro_sucursal, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(2, nro_sector, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(3, nro_categoria, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(4, nro_puesto, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(5, nro_ccosto, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(6, nro_gerencia, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(9, nro_departamento, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(10, nro_empresa, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(15, nro_cajajub, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(16, nro_sindicato, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(17, nro_osocial_elegida, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(18, nro_contrato, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(19, nro_convenio, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(20, nro_lpago, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(21, nro_reghorario, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(22, nro_formaliq, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(23, nro_planos_elegida, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(24, nro_osocial_ley, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(25, nro_planos_ley, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(29, nro_actividad, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(30, nro_sitrev, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(31, nro_condsijp, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(35, nro_direccion, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(40, nro_art, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(41, nro_bancopago, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(44, nro_locacion, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(45, nro_area, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(46, nro_SubDepto, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(47, nro_odon, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(48, nro_ose, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(49, nro_osl, NroTercero, F_Alta, F_Baja)
         
  If rs.State = adStateOpen Then
      rs.Close
  End If

End Sub

Public Sub LineaModelo_222()
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de DesmenFamiliar
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1            As Integer
Dim pos2            As Integer
Dim rs              As New ADODB.Recordset
Dim rsa             As New ADODB.Recordset
Dim Legajo          As Long ' Legajo del Empleado
Dim Anio            As String
Dim FecDesde        As String
Dim FecHasta        As String
Dim NroItem         As String
Dim Monto           As String
Dim NroTercero      As Integer
'Dim StrSql          As String

MyBeginTrans
  StrSql = " SELECT terfecnac,empleado,parenro FROM familiar "
  StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = familiar.ternro "
  StrSql = StrSql & " WHERE famcargadgi = -1 "
  OpenRecordset StrSql, rs
  
  Do While Not rs.EOF:
  
    NroTercero = rs!Empleado
    
    If rs!parenro = 1 Then
        NroItem = 10
    Else
        If rs!parenro = 2 Then
            NroItem = 11
        Else
            NroItem = 12
        End If
    End If
    If rs!terfecnac > CDate("01/01/2004") Then
        FecDesde = rs!terfecnac
    Else
        FecDesde = "01/01/2004"
    End If
    FecHasta = "31/12/2004"
    
    ' Inserto el Desmen
    StrSql = " SELECT desmondec FROM desmen WHERE empleado = " & NroTercero
    StrSql = StrSql & " AND itenro = " & NroItem
    StrSql = StrSql & " AND desano = 2004 "
    OpenRecordset StrSql, rsa
    
    If rsa.EOF Then
        ' Inserto el Desmen
        Monto = 1
        StrSql = " INSERT INTO desmen(empleado,itenro,desano,desfecdes,desfechas,desmenprorra,desmondec)"
        StrSql = StrSql & " VALUES(" & NroTercero & "," & NroItem & ",2004,'" & FecDesde & "','" & FecHasta & "',0," & Monto & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        Monto = Int(rsa!desmondec) + 1
        ' Actualizo el Desmen
        StrSql = " UPDATE desmen SET desmondec = " & Monto
        StrSql = StrSql & " WHERE empleado = " & NroTercero
        StrSql = StrSql & " AND itenro = " & NroItem
        StrSql = StrSql & " AND desano = 2004 "
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    rs.MoveNext
    
  Loop
  
  MyCommitTrans
  
  If rs.State = adStateOpen Then rs.Close
  If rsa.State = adStateOpen Then rsa.Close
End Sub

Public Sub LineaModelo_223(ByVal strReg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Legajo          As Long ' Legajo del Empleado
Dim Anio            As String
Dim FecDesde        As String
Dim FecHasta        As String
Dim NroItem         As String
Dim Monto           As String

Dim pos1            As Integer
Dim pos2            As Integer

Dim NroTercero      As Integer

Dim StrSql          As String
Dim rs              As New ADODB.Recordset

'    RegLeidos = RegLeidos + 1
    
    Flog.Writeline "Numero de Linea = " & RegLeidos
        
    pos1 = 1
    pos2 = InStr(pos1, strReg, Separador)
    Legajo = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Anio = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    FecDesde = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    FecHasta = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    NroItem = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = Len(strReg) + 1
    Monto = Mid(strReg, pos1, pos2 - pos1)
    
' Busco el empleado

  StrSql = " SELECT ternro FROM empleado WHERE empleg = " & Legajo
  OpenRecordset StrSql, rs
  NroTercero = rs!ternro

  Flog.Writeline "Legajo = " & Legajo & "Codigo de Tercero = " & NroTercero

' Inserto el Desmen
  StrSql = " INSERT INTO desmen(empleado,itenro,desano,desfecdes,desfechas,desmenprorra,desmondec)"
  StrSql = StrSql & " VALUES(" & NroTercero & "," & NroItem & "," & Anio & ",'" & FecDesde & "','" & FecHasta & "',0," & Monto & ")"
  objConn.Execute StrSql, , adExecuteNoRecords
  Flog.Writeline "Inserte el item - " & NroItem & " - " & Anio & " - " & Monto
  
  If rs.State = adStateOpen Then rs.Close
End Sub

Public Sub LineaModelo_224(ByVal strReg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Legajo          As Long ' Legajo del Empleado
Dim Anio            As String  ' Apellido del Familiar
Dim mes             As String  ' Nombre del Familiar
Dim Item1           As String  ' Item 1
Dim Item2           As String  ' Item 2
Dim Item3           As String  ' Item 3
Dim Item4           As String  ' Item 4
Dim Item5           As String  ' Item 5
Dim Item6           As String  ' Item 6
Dim Item7           As String  ' Item 7
Dim Item8           As String  ' Item 8
Dim Item9           As String  ' Item 9
Dim Item10           As String  ' Item 10
Dim Item11           As String  ' Item 11
Dim Item12           As String  ' Item 12
Dim Item13           As String  ' Item 13
Dim Item14           As String  ' Item 14
Dim Item15           As String  ' Item 15
Dim Item16           As String  ' Item 16
Dim Item17           As String  ' Item 17
Dim Item18           As String  ' Item 18
Dim Item19           As String  ' Item 19
Dim Item20           As String  ' Item 20
Dim Item21           As String  ' Item 21
Dim Item22           As String  ' Item 22

Dim pos1            As Integer
Dim pos2            As Integer

Dim NroTercero      As Integer
Dim FecHasta_Peri   As String

Dim StrSql          As String
Dim rs              As New ADODB.Recordset

    Item1 = ""
    Item2 = ""
    Item3 = ""
    Item4 = ""
    Item5 = ""
    Item6 = ""
    Item7 = ""
    Item8 = ""
    Item9 = ""
    Item10 = ""
    Item11 = ""
    Item12 = ""
    Item13 = ""
    Item14 = ""
    Item15 = ""
    Item16 = ""
    Item17 = ""
    Item18 = ""
    Item19 = ""
    Item20 = ""
    Item21 = ""
    Item22 = ""
    
'    RegLeidos = RegLeidos + 1
    
    Flog.Writeline "Numero de Linea = " & RegLeidos
        
    pos1 = 1
    pos2 = InStr(pos1, strReg, Separador)
    Legajo = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Anio = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    mes = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item1 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item2 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item3 = Mid(strReg, pos1, pos2 - pos1)
    
    'pos1 = pos2 + 1
    'pos2 = InStr(pos1 + 1, strReg, separador)
    'Item4 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item5 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item6 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item7 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item8 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item9 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item10 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item11 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item12 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item13 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item14 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item15 = Mid(strReg, pos1, pos2 - pos1)
    
    'pos1 = pos2 + 1
    'pos2 = InStr(pos1 + 1, strReg, separador)
    'Item16 = Mid(strReg, pos1, pos2 - pos1)
    
    'pos1 = pos2 + 1
    'pos2 = InStr(pos1 + 1, strReg, separador)
    'Item17 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item18 = Mid(strReg, pos1, pos2 - pos1)
    
    'pos1 = pos2 + 1
    'pos2 = InStr(pos1 + 1, strReg, separador)
    'Item19 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item20 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item22 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = Len(strReg) + 1
    Item21 = Mid(strReg, pos1, pos2 - pos1)
    
    If (mes = 1) Then
        FecHasta_Peri = "31/01/" & Anio
    Else
        If (mes = 2) Then
            FecHasta_Peri = "28/02/" & Anio
        Else
            If (mes = 3) Then
                FecHasta_Peri = "31/03/" & Anio
            Else
                If (mes = 4) Then
                    FecHasta_Peri = "30/04/" & Anio
                Else
                    If (mes = 5) Then
                        FecHasta_Peri = "31/05/" & Anio
                    Else
                        If (mes = 6) Then
                            FecHasta_Peri = "30/06/" & Anio
                        Else
                            If (mes = 7) Then
                                FecHasta_Peri = "31/07/" & Anio
                            Else
                                If (mes = 8) Then
                                    FecHasta_Peri = "31/08/" & Anio
                                Else
                                    If (mes = 9) Then
                                        FecHasta_Peri = "30/09/" & Anio
                                    Else
                                        If (mes = 10) Then
                                            FecHasta_Peri = "31/10/" & Anio
                                        Else
                                            If (mes = 11) Then
                                                FecHasta_Peri = "30/11/" & Anio
                                            Else
                                                If (mes = 12) Then
                                                    FecHasta_Peri = "31/12/" & Anio
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
' Busco el empleado

  StrSql = " SELECT ternro FROM empleado WHERE empleg = " & Legajo
  OpenRecordset StrSql, rs
  NroTercero = rs!ternro

  Flog.Writeline "Legajo = " & Legajo & "Codigo de Tercero = " & NroTercero

' Inserto el Desliq item 1
  If Item1 <> "0" And Item1 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",1,'" & FecHasta_Peri & "',Null," & Item1 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item1 - " & Legajo
  End If
' Inserto el Desliq item 2
  If Item2 <> "0" And Item2 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",2,'" & FecHasta_Peri & "',Null," & Item2 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item2 - " & Legajo
  End If
' Inserto el Desliq item 3
  If Item3 <> "0" And Item3 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",3,'" & FecHasta_Peri & "',Null," & Item3 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item3 - " & Legajo
  End If
' Inserto el Desliq item 4
  If Item4 <> "0" And Item4 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",4,'" & FecHasta_Peri & "',Null," & Item4 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item4 - " & Legajo
  End If
' Inserto el Desliq item 5
  If Item5 <> "0" And Item5 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",5,'" & FecHasta_Peri & "',Null,-" & Item5 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item5 - " & Legajo
  End If
' Inserto el Desliq item 6
  If Item6 <> "0" And Item6 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",6,'" & FecHasta_Peri & "',Null,-" & Item6 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item6 - " & Legajo
  End If
' Inserto el Desliq item 7
  If Item7 <> "0" And Item7 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",7,'" & FecHasta_Peri & "',Null," & Item7 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item7 - " & Legajo
  End If
' Inserto el Desliq item 8
  If Item8 <> "0" And Item8 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",8,'" & FecHasta_Peri & "',Null," & Item8 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item8 - " & Legajo
  End If
' Inserto el Desliq item 9
  If Item9 <> "0" And Item9 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",9,'" & FecHasta_Peri & "',Null," & Item9 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item9 - " & Legajo
  End If
' Inserto el Desliq item 10
  If Item10 <> "0" And Item10 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",10,'" & FecHasta_Peri & "',Null," & Item10 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item10 - " & Legajo
  End If
' Inserto el Desliq item 11
  If Item11 <> "0" And Item11 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",11,'" & FecHasta_Peri & "',Null," & Item11 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item11 - " & Legajo
  End If
' Inserto el Desliq item 12
  If Item12 <> "0" And Item12 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",12,'" & FecHasta_Peri & "',Null," & Item12 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item12 - " & Legajo
  End If
' Inserto el Desliq item 13
  If Item13 <> "0" And Item13 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",13,'" & FecHasta_Peri & "',Null," & Item13 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item13 - " & Legajo
  End If
' Inserto el Desliq item 14
  If Item14 <> "0" And Item14 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",14,'" & FecHasta_Peri & "',Null," & Item14 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item14 - " & Legajo
  End If
' Inserto el Desliq item 15
  If Item15 <> "0" And Item15 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",15,'" & FecHasta_Peri & "',Null," & Item15 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item15 - " & Legajo
  End If
' Inserto el Desliq item 16
  If Item16 <> "0" And Item16 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",16,'" & FecHasta_Peri & "',Null," & Item16 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item16 - " & Legajo
  End If
' Inserto el Desliq item 17
  If Item17 <> "0" And Item17 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",17,'" & FecHasta_Peri & "',Null," & Item17 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item17 - " & Legajo
  End If
' Inserto el Desliq item 18
  If Item18 <> "0" And Item18 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",18,'" & FecHasta_Peri & "',Null," & Item18 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item18 - " & Legajo
  End If
' Inserto el Desliq item 19
  If Item19 <> "0" And Item19 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",19,'" & FecHasta_Peri & "',Null," & Item19 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item19 - " & Legajo
  End If
' Inserto el Desliq item 20
  If Item20 <> "0" And Item20 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",20,'" & FecHasta_Peri & "',Null," & Item20 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item20 - " & Legajo
  End If
' Inserto el Desliq item 21
  If Item21 <> "0" And Item21 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",21,'" & FecHasta_Peri & "',Null," & Item21 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item21 - " & Legajo
  End If
' Inserto el Desliq item 22
  If Item22 <> "0" And Item22 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",22,'" & FecHasta_Peri & "',Null," & Item22 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item22 - " & Legajo
  End If
  
  If rs.State = adStateOpen Then rs.Close

End Sub


Public Sub LineaModelo_225(ByVal strReg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1            As Integer
Dim pos2            As Integer
Dim Legajo          As String   'Legajo                   -- empleado.empleg
Dim Anio            As String   'Año de la liquidacion    -- periodo.pliqanio
Dim mes             As String   'Mes de la liquidacion    -- periodo.pliqmes
Dim Proceso         As String   'Proceso de liquidacion   -- proceso.prodesc
Dim CtoCodigo       As String   'Código de concepto       -- concepto.conccod
Dim Monto           As String   'Monto liquidado          -- detliq.dlimonto
Dim Cantidad        As String   'Cantidad liquidada       -- detliq.dlicant

Dim Desc_Periodo    As String   'Descripcion del Periodo de liquidacion
Dim FecDesde_Peri   As String   'Fecha desde del periodo
Dim FecHasta_Peri   As String   'Fecha desde del periodo

Dim NroTercero          As Integer

Dim Nro_Legajo          As Integer
Dim nro_concepto        As Integer
Dim nro_periodo         As Integer
Dim nro_proceso         As Integer
Dim nro_cabecera        As Integer
Dim nro_tipoconc        As Integer

Dim RsPeriodo    As New ADODB.Recordset
Dim RsPeri       As New ADODB.Recordset
Dim RsConcepto   As New ADODB.Recordset
Dim RsCabecera   As New ADODB.Recordset
Dim RsCabe       As New ADODB.Recordset
Dim RsCabliq     As New ADODB.Recordset
Dim RsProceso    As New ADODB.Recordset
Dim RsPro        As New ADODB.Recordset
Dim RsEmple      As New ADODB.Recordset

Dim CodPeri As Integer
Dim CodPro  As Integer
Dim CodCabe As Integer

'    RegLeidos = RegLeidos + 1
    
    Flog.Writeline "Numero de Linea = " & RegLeidos
    
    ' Recupero los Valores del Archivo
    
    pos1 = 1
    pos2 = InStr(pos1, strReg, Separador)
    Legajo = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Anio = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    mes = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Proceso = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    CtoCodigo = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Cantidad = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = Len(strReg) + 1
    Monto = Mid(strReg, pos1, pos2 - pos1)
    
' Busco al empleado asociado

  StrSql = " SELECT ternro,empleg FROM empleado WHERE empleg = " & Legajo
  OpenRecordset StrSql, RsEmple
  
  If RsEmple.EOF Then 'No existe el empleado
    NroTercero = 0
  Else
    NroTercero = RsEmple!ternro
  End If
  
  If NroTercero <> 0 Then
      Flog.Writeline "Procesando al empleado = " & RsEmple!empleg
    
    ' Busco el periodo de liquidacion
    
      StrSql = " SELECT pliqnro FROM periodo WHERE pliqanio = " & Anio
      StrSql = StrSql & " AND pliqmes = " & mes
      OpenRecordset StrSql, RsPeriodo
      
      If RsPeriodo.EOF Then  'No existe el periodo => lo creo
        If (mes = 1) Then
            Desc_Periodo = "Enero "
            FecDesde_Peri = "01/01/" & Anio
            FecHasta_Peri = "31/01/" & Anio
        Else
            If (mes = 2) Then
                Desc_Periodo = "Febrero "
                FecDesde_Peri = "01/02/" & Anio
                FecHasta_Peri = "28/02/" & Anio
            Else
                If (mes = 3) Then
                    Desc_Periodo = "Marzo "
                    FecDesde_Peri = "01/03/" & Anio
                    FecHasta_Peri = "31/03/" & Anio
                Else
                    If (mes = 4) Then
                        Desc_Periodo = "Abril "
                        FecDesde_Peri = "01/04/" & Anio
                        FecHasta_Peri = "30/04/" & Anio
                    Else
                        If (mes = 5) Then
                            Desc_Periodo = "Mayo "
                            FecDesde_Peri = "01/05/" & Anio
                            FecHasta_Peri = "31/05/" & Anio
                        Else
                            If (mes = 6) Then
                                Desc_Periodo = "Junio "
                                FecDesde_Peri = "01/06/" & Anio
                                FecHasta_Peri = "30/06/" & Anio
                            Else
                                If (mes = 7) Then
                                    Desc_Periodo = "Julio "
                                    FecDesde_Peri = "01/07/" & Anio
                                    FecHasta_Peri = "31/07/" & Anio
                                Else
                                    If (mes = 8) Then
                                        Desc_Periodo = "Agosto "
                                        FecDesde_Peri = "01/08/" & Anio
                                        FecHasta_Peri = "31/08/" & Anio
                                    Else
                                        If (mes = 9) Then
                                            Desc_Periodo = "Septiembre "
                                            FecDesde_Peri = "01/09/" & Anio
                                            FecHasta_Peri = "30/09/" & Anio
                                        Else
                                            If (mes = 10) Then
                                                Desc_Periodo = "Octubre "
                                                FecDesde_Peri = "01/10/" & Anio
                                                FecHasta_Peri = "31/10/" & Anio
                                            Else
                                                If (mes = 11) Then
                                                    Desc_Periodo = "Noviembre "
                                                    FecDesde_Peri = "01/11/" & Anio
                                                    FecHasta_Peri = "30/11/" & Anio
                                                Else
                                                    If (mes = 12) Then
                                                        Desc_Periodo = "Diciembre "
                                                        FecDesde_Peri = "01/12/" & Anio
                                                        FecHasta_Peri = "31/12/" & Anio
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Desc_Periodo = Desc_Periodo & Anio
        
        StrSql = " INSERT INTO periodo(pliqmes,pliqanio,pliqdesc,pliqdesde,pliqhasta) "
        StrSql = StrSql & " VALUES(" & mes & "," & Anio & ",'" & Desc_Periodo & "','" & FecDesde_Peri & "','" & FecHasta_Peri & "')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = " SELECT @@IDENTITY AS CodPeri "
        OpenRecordset StrSql, RsPeri
        
        nro_periodo = RsPeri!CodPeri
        
      Else
        nro_periodo = RsPeriodo!PliqNro
      End If
      
    ' Busco el proceso dentro del periodo de liquidacion
    
      StrSql = " SELECT pronro FROM proceso WHERE prodesc = '" & Proceso & "'"
      StrSql = StrSql & " AND pliqnro = " & nro_periodo
      OpenRecordset StrSql, RsProceso
      
      If RsProceso.EOF Then  'No existe el proceso => lo creo
        StrSql = " INSERT INTO proceso(pliqnro,tprocnro,profecpago,prodesc,profecplan,profecini,profecfin) "
        StrSql = StrSql & " VALUES(" & nro_periodo & ",3,'" & FecHasta_Peri & "','" & Proceso & "','"
        StrSql = StrSql & FecHasta_Peri & "','" & FecDesde_Peri & "','" & FecHasta_Peri & "')"
        objConn.Execute StrSql, , adExecuteNoRecords
        StrSql = " SELECT @@IDENTITY AS CodPro "
        OpenRecordset StrSql, RsPro
        nro_proceso = RsPro!CodPro
      Else
        nro_proceso = RsProceso!pronro
      End If
      
    ' Busco el concepto
    
      StrSql = " SELECT concnro,tconnro FROM concepto WHERE conccod = '" & CtoCodigo & "'"
      OpenRecordset StrSql, RsConcepto
      
      If RsConcepto.EOF Then  'No existe el concepto => Error
        nro_concepto = 0
      Else
        nro_concepto = RsConcepto!concnro
'Ajuste para Accor porque los conceptos son todos positivos -------------------------------------
        nro_tipoconc = RsConcepto!tconnro
        If nro_tipoconc = 6 Or nro_tipoconc = 8 Or nro_tipoconc = 10 Or nro_tipoconc = 13 Then
            If Mid(Monto, 1, 1) = "-" Then
                Monto = Mid(Monto, 2)
            Else
                Monto = "-" & Monto
            End If
        End If
'Fin ajuste para Accor --------------------------------------------------------------------------
      End If
    
    ' Busco el cabliq del empleado para el proceso y periodo
    
      StrSql = " SELECT cliqnro FROM cabliq WHERE empleado = " & NroTercero
      StrSql = StrSql & " AND pronro = " & nro_proceso
      OpenRecordset StrSql, RsCabliq
      
      If RsCabliq.EOF Then  'No existe el cabliq => lo creo
        StrSql = " INSERT INTO cabliq(empleado,pronro) VALUES("
        StrSql = StrSql & NroTercero & "," & nro_proceso & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = " SELECT @@IDENTITY AS CodCabe "
        OpenRecordset StrSql, RsCabe
        
        nro_cabecera = RsCabe!CodCabe
      Else
        nro_cabecera = RsCabliq!cliqnro
      End If
      
      If nro_concepto <> 0 Then ' Inserto el detalle de liquidacion
      
        StrSql = " INSERT INTO detliq(cliqnro,concnro,dlimonto,dlicant,dliqdesde,dliqhasta,tconnro,dlitexto,dlifec) VALUES("
        StrSql = StrSql & nro_cabecera & "," & nro_concepto & "," & Monto & "," & Cantidad
        StrSql = StrSql & ",0,0,0,'','00:00:00')"
        objConn.Execute StrSql, , adExecuteNoRecords
    
        Flog.Writeline "Inserte el Detalle de liquidacion  - " & CtoCodigo & " - " & mes & " - " & Anio
      Else
        Flog.Writeline "Concepto Inexistente = " & CtoCodigo
      End If
  Else
      Flog.Writeline "Empleado Inexistente = " & RsEmple!empleg
  End If
End Sub


'Public Sub LineaModelo_226(ByVal strReg As String)
'' ---------------------------------------------------------------------------------------------
'' Descripcion: Interface de Postulantes
'' Autor      : FGZ
'' Fecha      : 09/08/2004
'' Ultima Mod.: 12/10/2004 - Lisandro Moro
''   Descripcion:  Se actualiza y agregan campos al estandar.
''   INCLUYE:    TraerCodTipoDocumento()
''               TraerCodLocalidad()
''               TraerCodProvincia()
''               TraerCodZona()
''               TraerCodPais()
''               TraerCodTitulo()
'' ---------------------------------------------------------------------------------------------
'Dim pos1            As Integer
'Dim pos2            As Integer
'Dim aux             As String
'
''-------------------------------
''El formato de los campos es:
''TITULO(Majusculas)
'    'campos
''-------------------------------
''POSTULANTE
'    '(01)Tipo Doc
'    '(02)nro Doc
'    '(03)nombre 1
'    '(04)nombre 2
'    '(05)Apellido 1
'    '(06)Apellido 2
''TELEFONOS
'    '(07)Telefono Personal
'    '(08)Telefono para mensajes
''POSTULANTE
'    '(09)Email
''DOMICILIO
'    '(10)Dir: Calle
'    '(11)Dir: nro
'    '(12)Dir: Piso
'    '(13)Dir: Dpto
'    '(14)C.P
'    '(15)Localidad
'    '(16)Provincia
'    '(17)Partido
'    '(18)Zona
'    '(19)Pais
''POSTULANTE
'    '(20)Naciemiento
'    '(21)Sexo
''ESTUDIOS FORMALES 1
'    '(22)Nivel
'    '(23)Completo
'    '(24)Titulo
'    '(25)Institucion
''ESTUDIOS FORMALES1 2
'    '(26)Posgrado
'    '(27)Postitulo
'    '(28)Institucion
''EXPERIENCIA LABORAL 1
'    '(29)Cargo anterior
'    '(30)Empresa
'    '(31)Tarea Desempeñada
'    '(32)Fec Desde
'    '(33)Fec Hasta
''EXPERIENCIA LABORAL 2
'    '(34)Cargo anterior 2
'    '(35)Empresa
'    '(36)Tarea Desempeñada
'    '(37)Fec Desde
'    '(38)Fec Hasta
''IDIOMAS
'    '(39)Idioma 1
'    '(40)Lee: Nivel
'    '(41)Habla: Nivel
'    '(42)Escribe: Nivel
'    '(43)Idioma 2
'    '(44)Lee: Nivel
'    '(45)Habla: Nivel
'    '(46)Escribe: Nivel
'    '(47)Idioma 3
'    '(48)Lee: Nivel
'    '(49)Habla: Nivel
'    '(50)Escribe: Nivel
''OTROS CURSOS 1
'    '(51)desc.Curso
'    '(52)Tipo Curso
'    '(53)Fec Curso
'    '(54)Institucion
''OTROS CURSOS 2
'    '(55)desc.Curso
'    '(56)Tipo Curso
'    '(57)Fec Curso
'    '(58)Institucion
'
'' El Formato era
'                'NroDoc,
'                'Nombre,
'                'Apellido,
'                'Nro Telefono(Default),
'                'Celular,
'                'Email,
'                'Direccion,
'                'CP,
'                'Localidad,
'                'Provincia,
'                'Fecha Nac,
'                'Sexo,
'                'Nivel Estudio,
'                'Estudio Completo,
'                'Titulo,
'                'Institucion,
'                'Postgrado,
'                'Postitulo,
'                'Cargo,
'' por ahora no                        'Cargo Ejecutivo,
'' por ahora no                        'Experiencia,
'' por ahora no                        'Años,
'' por ahora no                        'Datos, ...
'
'
''POSTULANTE
'    Dim tidnro As Integer       '(01)Tipo Doc
'    Dim nrodoc As String        '(02)nro Doc
'    Dim ternom As String        '(03)nombre 1
'    Dim ternom2 As String       '(04)nombre 2
'    Dim terape As String        '(05)Apellido 1
'    Dim terape2 As String       '(06)Apellido 2
''TELEFONOS
'    Dim telnro As String        '(07)Telefono Personal
'    Dim telnro2 As String       '(08)Telefono para mensajes
''POSTULANTE
'    Dim teremail As String      '(09)Email
''DOMICILIO
'    Dim calle As String         '(10)Dir: Calle
'    Dim nro As String           '(11)Dir: nro
'    Dim piso As String          '(12)Dir: Piso
'    Dim oficdepto As String     '(13)Dir: Dpto
'    Dim codigipostal As String  '(14)C.P
'    Dim locnro As Integer       '(15)Localidad
'    Dim provnro As Integer      '(16)Provincia
'    Dim partnro As Integer      '(17)Partido
'    Dim zonanro As Integer      '(18)Zona
'    Dim paisnro As Integer      '(19)Pais
''POSTULANTE
'    Dim terfecnac As Date       '(20)Naciemiento
'    Dim tersex As Boolean       '(21)Sexo
''ESTUDIOS FORMALES 1
'    Dim Nivnro As Integer       '(22)Nivel
'    Dim capcomp As Boolean      '(23)Completo
'    Dim titnro As Integer       '(24)Titulo
'    Dim instnro As Integer      '(25)Institucion
''ESTUDIOS FORMALES1 2
'    Dim nivnro2 As Integer      '(26)Posgrado
'    Dim titnro2 As Integer      '(27)Postitulo
'    Dim instnro2 As Integer     '(28)Institucion
''EXPERIENCIA LABORAL 1
'    Dim carnro As Integer       '(29)Cargo anterior
'    Dim lempnro As Integer      '(30)Empresa
'    Dim empatareas As String    '(31)Tarea Desempeñada
'    Dim empaini As Date         '(32)Fec Desde
'    Dim empafin As Date         '(33)Fec Hasta
''EXPERIENCIA LABORAL 2
'    Dim carnro2 As Integer      '(34)Cargo anterior 2
'    Dim lempnro2 As Integer     '(35)Empresa
'    Dim empatareas2 As String   '(36)Tarea Desempeñada
'    Dim empaini2 As Date        '(37)Fec Desde
'    Dim empafin2 As Date        '(38)Fec Hasta
''IDIOMAS
'    Dim idinro1 As Integer      '(39)Idioma 1
'    Dim emidlee1 As Integer     '(40)Lee: Nivel
'    Dim empidhabla1 As Integer  '(41)Habla: Nivel
'    Dim empidescr1 As Integer   '(42)Escribe: Nivel
'    Dim idinro2 As Integer      '(43)Idioma 2
'    Dim emidlee2 As Integer     '(44)Lee: Nivel
'    Dim empidhabla2 As Integer  '(45)Habla: Nivel
'    Dim empidescr2 As Integer   '(46)Escribe: Nivel
'    Dim idinro3 As Integer      '(47)Idioma 3
'    Dim emidlee3 As Integer     '(48)Lee: Nivel
'    Dim empidhabla3 As Integer  '(49)Habla: Nivel
'    Dim empidescr3 As Integer   '(50)Escribe: Nivel
''OTROS CURSOS 1
'    Dim estinfdesabr As String  '(51)desc.Curso
'    Dim tipcurnro As Integer    '(52)Tipo Curso
'    Dim estinffecha As Date     '(53)Fec Curso
'    Dim instnro3 As Integer     '(54)Institucion
''OTROS CURSOS 2
'    Dim estinfdesabr2 As String '(55)desc.Curso
'    Dim tipcurnro2 As Integer   '(56)Tipo Curso
'    Dim estinffecha2 As Date    '(57)Fec Curso
'    'Dim instnro3 As Integer     '(58)Institucion
'
'    Dim ternro As Integer
'    'Dim ArrStr
'    'ArrStr =
'
'Dim Ter_Apellido        As String
'Dim Ter_Nombre          As String
'Dim Fecha_Nacimiento    As String
'Dim Sexo                As String
'Dim NivEst              As String
'Dim Docu_Nro            As String
'Dim Docu_Tipo           As String
'Dim Dir_Calle           As String
'Dim Dir_Nro             As String
'Dim Dir_Piso            As String
'Dim Dir_OficDepto       As String
'Dim Dir_Sector          As String
'Dim Dir_Torre           As String
'Dim Dir_Manzana         As String
'Dim Dir_CP              As String
'Dim Dir_Localidad       As String
'Dim Localidad_Nro       As String
'Dim Dir_Provincia       As String
'Dim Provincia_Nro       As String
'Dim Dir_Pais            As String
'Dim Pais_Nro            As String
'Dim Tel_Default         As String
'Dim Tel_Celular         As String
'Dim Email               As String
'Dim NroTercero          As Integer
'
'Dim Ne_Descripcion          As String
'Dim Ne_Nro                  As String
'Dim Ne_Completo             As Boolean
'Dim Titulo_Descripcion      As String
'Dim Titulo_Nro              As String
'Dim Institucion_Descripcion As String
'Dim Institucion_Nro         As String
'Dim PosTitulo_Descripcion   As String
'Dim PosTitulo_Nro           As String
'Dim tiene_postgrado         As Boolean
'Dim Cargo_Descripcion       As String
'Dim Cargo_Nro               As String
'
'Dim NroDom                  As Long
'Dim ArrStr
'
'Dim rs As New ADODB.Recordset
'Dim rs_sql As New ADODB.Recordset
'
'
'    'Cargo los datos le la fila en un arreglo
'    ArrStr = Split(strReg, Separador)
'
'    'MsgBox (ArrStr(5))
'    'Recupero los Valores del Archivo
'
'    'Columna 0 - Tipo de Documento
'    If Not EsNulo(ArrStr(0)) Then
'        'por defecto tomo DNI si el valor es vacio
'        tidnro = 1
'    Else
'        tidnro = TraerCodTipoDocumento(Trim(ArrStr(0)))
'    End If
'
'    'Columna 1 - Nro de documento
'    If Not EsNulo(ArrStr(1)) Or IsNumeric(CInt(ArrStr(1))) Then
'        nrodoc = Left(Trim(ArrStr(1)), 30)
'    Else
'        nrodoc = ""
'    End If
'
'    'Columna 2 - Apellido(terape)
'    If Not EsNulo(ArrStr(2)) Then
'        terape = Left(Trim(ArrStr(2)), 25)
'    Else
'        terape = ""
'    End If
'
'    'Columna 3 - Apellido(terape2)
'    If Not EsNulo(ArrStr(3)) Then
'        terape2 = Left(Trim(ArrStr(3)), 25)
'    Else
'        terape2 = ""
'    End If
'
'    'Columna 4 - Apellido(ternom)
'    If Not EsNulo(ArrStr(4)) Then
'        ternom = Left(Trim(ArrStr(4)), 25)
'    Else
'        ternom = ""
'    End If
'
'    'Columna 5 - Apellido(terape2)
'    If Not EsNulo(ArrStr(5)) Then
'        ternom2 = Left(Trim(ArrStr(5)), 25)
'    Else
'        ternom2 = "Null"
'    End If
'
'    'Columna 6 - Telefono Personal
'    If Not EsNulo(ArrStr(6)) Then
'        telnro = Left(Trim(ArrStr(6)), 20)
'    Else
'        telnro = "Null"
'    End If
'
'    'Columna 7 - Telefono Mensaje
'    If Not EsNulo(ArrStr(7)) Then
'        telnro2 = Left(Trim(ArrStr(7)), 20)
'    Else
'        telnro2 = "Null"
'    End If
'
'    'Columna 8 - eMail
'    If Not EsNulo(ArrStr(8)) Then
'        teremail = Left(Trim(ArrStr(8)), 20)
'    Else
'        teremail = "Null"
'    End If
'
'    'Columna  9 - Direccion(calle)
'    If Not EsNulo(ArrStr(9)) Then
'        calle = Left(Trim(ArrStr(9)), 30)
'    Else
'        calle = "Null"
'    End If
'
'    'Columna  10 - Direccion(nro)
'    If Not EsNulo(ArrStr(10)) Then
'        nro = Left(Trim(ArrStr(10)), 8)
'    Else
'        nro = "Null"
'    End If
'
'    'Columna  11 - Direccion(Piso)
'    If Not EsNulo(ArrStr(11)) Then
'        piso = Left(Trim(ArrStr(11)), 8)
'    Else
'        piso = "Null"
'    End If
'
'    'Columna  12 - Direccion(Piso)
'    If Not EsNulo(ArrStr(12)) Then
'        oficdepto = Left(Trim(ArrStr(12)), 8)
'    Else
'        oficdepto = "Null"
'    End If
'
'    'Columna  13 - Direccion(C.P.)
'    If Not EsNulo(ArrStr(13)) Then
'        codigopostal = Left(Trim(ArrStr(13)), 12)
'    Else
'        codigopostal = "Null"
'    End If
'
'    'Columna  14 - Direccion(localidad)
'    If Not EsNulo(ArrStr(14)) Then
'        locnro = CInt(TraerCodLocalidad(Left(Trim(ArrStr(14)), 30)))
'    Else
'        locnro = 1 'No Informada
'    End If
'
'    'Columna  15 - Direccion(Provincia)
'    If Not EsNulo(ArrStr(15)) Then
'        locnro = CInt(TraerCodProvincia(Left(Trim(ArrStr(15)), 20)))
'    Else
'        locnro = 1 'No Informada
'    End If
'
'    'Columna  16 - Direccion(Partido)
'    If Not EsNulo(ArrStr(16)) Then
'        partnro = CInt(TraerCodPartido(Left(Trim(ArrStr(16)), 20)))
'    Else
'        partnro = 1 'Sin datos
'    End If
'
'    'Columna  17 - Direccion(Zona)
'    If Not EsNulo(ArrStr(17)) Then
'        zonanro = CInt(TraerCodZona(Left(Trim(ArrStr(17)), 20)))
'    Else
'        zonanro = 1 'Capital
'    End If
'
'    'Columna  18 - Direccion(Pais)
'    If Not EsNulo(ArrStr(18)) Then
'        zonanro = CInt(TraerCodPais(Left(Trim(ArrStr(18)), 20)))
'    Else
'        zonanro = 1 'No informado
'    End If
'
'    'Columna 19 - Fecha Nacimiento
'    ArrStr(19) = Left(Trim(ArrStr(19)), 20)
'    If UCase(ArrStr(19)) = "NULL" Then
'       terfecnac = "Null" 'No informado
'    Else
'        If IsDate(ArrStr(19)) Then
'            terfecnac = ConvFecha(ArrStr(19))
'        Else
'            Flog.Writeline "Fecha de Nacimiento invalida " & ArrStr(19)
'            InsertaError 11, 4
'            terfecnac = "Null" 'No informado
'        End If
'    End If
'
'
'    'Columna 20 - Sexo
'    ArrStr(20) = UCase(Trim(ArrStr(19)))
'    If (ArrStr(20) = "MUJER") Or (ArrStr(20) = "M") Or (ArrStr(20) = "F") Or (ArrStr(20) = "FEMENINO") Then
'        tersex = False
'    Else
'        tersex = True
'    End If
'
'    'Columna 21 - Nivel de estudio.(Descripcion nivel)
'    If Not EsNulo(ArrStr(21)) Then
'        Nivnro = CInt(TraerCodNivelEstudio(Left(Trim(ArrStr(21)), 40)))
'    Else
'        Nivnro = 5 'Sin datos
'    End If
'
'    'Columna 22 - Nivel de estudio.(Completo)
'    ArrStr(22) = UCase(Trim(ArrStr(22)))
'    If (ArrStr(22) = "SI") Or (ArrStr(22) = "TRUE") Or (ArrStr(22) = "-1") Or (ArrStr(22) = "COMPLETO") Then
'        capcomp = True
'    Else
'        capcomp = False
'    End If
'
'    'Columna 23 - Nivel de estudio.(Titulo)
'    If Not EsNulo(ArrStr(23)) Then
'        titnro = CInt(TraerCodTitulo(Left(Trim(ArrStr(23)), 40), Nivnro))
'    Else
'        titnro = 0 'Sin datos
'    End If
'
'    'Columna 24 - Nivel de estudio.(Institucion)
'    If Not EsNulo(ArrStr(24)) Then
'        instnro = CInt(TraerCodInstitucion(Left(Trim(ArrStr(24)), 200)))
'    Else
'        instnro = 7 'No informada
'    End If
'
'    'Columna 25 - Nivel de estudio.(Posgrados Nivel)
'    If Not EsNulo(ArrStr(25)) Then
'        nivnro2 = CInt(TraerCodNivelEstudio(Left(Trim(ArrStr(25)), 40)))
'    Else
'        nivnro2 = 5 'Sin datos
'    End If
'
'    'Columna 26 - Nivel de estudio.(Posgrado Titulo)
'    If Not EsNulo(ArrStr(26)) Then
'        titnro2 = CInt(TraerCodTitulo(Left(Trim(ArrStr(26)), 40), Nivnro))
'    Else
'        titnro2 = 0 'Sin datos
'    End If
'
'    'Columna 27 - Nivel de estudio.(Posgrado Institucion)
'    If Not EsNulo(ArrStr(27)) Then
'        instnro2 = CInt(TraerCodInstitucion(Left(Trim(ArrStr(27)), 200)))
'    Else
'        instnro2 = 7 'No informada
'    End If
'
'    'Columna 28 - Experiencia laboral(Cargo)
'    carnro = CInt(TraerCodCargo(Left(Trim(ArrStr(28)), 50)))
'
'    'Columna 29 - Experiencia laboral(Empresa)(60)
'
'    lempnro = a
'    If Left(Trim(ArrStr(28)), 50) <> "N/A" Then
'
'            Call ValidaEstructuraCodExt(10, Empresa, nro_empresa, Inserto_estr)
'
'    Else
'        nro_empresa = 0
'    End If
'
'
'
'
'
'    If Not IsNull(Institucion_Descripcion) Then
'        'Valido que exista la institucion. Si no existe lo creo
'        StrSql = " SELECT * FROM institucion "
'        StrSql = StrSql & " WHERE instdes = '" & Institucion_Descripcion & "'"
'        If rs_sql.State = adStateOpen Then rs_sql.Close
'        OpenRecordset StrSql, rs_sql
'        If rs_sql.EOF Then
'            'Lo creo
'            StrSql = " INSERT INTO institucion (instdes,instnro,instabre) VALUES ("
'            StrSql = StrSql & "'" & Titulo_Descripcion & "'"
'            StrSql = StrSql & ",'" & Left(Titulo_Descripcion, 30) & "'"
'            StrSql = StrSql & ")"
'            objConn.Execute StrSql, , adExecuteNoRecords
'
'            Institucion_Nro = getLastIdentity(objConn, "institucion")
'        Else
'            Institucion_Nro = rs_sql!instnro
'        End If
'    Else
'        'Valido que exista un registro N\A. Si no existe lo creo
'        StrSql = " SELECT * FROM institucion "
'        StrSql = StrSql & " WHERE instdes = 'Sin Datos'"
'        If rs_sql.State = adStateOpen Then rs_sql.Close
'        OpenRecordset StrSql, rs_sql
'        If rs_sql.EOF Then
'            'Lo creo
'            StrSql = " INSERT INTO institucion (instdes,instnro,instabre) VALUES ("
'            StrSql = StrSql & "'Sin Datos'"
'            StrSql = StrSql & ",'Sin Datos'"
'            StrSql = StrSql & ")"
'            objConn.Execute StrSql, , adExecuteNoRecords
'
'            Institucion_Nro = getLastIdentity(objConn, "institucion")
'        Else
'            Institucion_Nro = rs_sql!instnro
'        End If
'    End If
'
'    If Not IsNull(Titulo_Descripcion) Then
'        'Valido que exista el titulo. Si no existe lo creo
'        StrSql = " SELECT * FROM titulo "
'        StrSql = StrSql & " WHERE titdesabr = '" & Titulo_Descripcion & "'"
'        StrSql = StrSql & " AND instnro =" & Institucion_Nro
'        StrSql = StrSql & " AND nivnro =" & Ne_Nro
'        If rs_sql.State = adStateOpen Then rs_sql.Close
'        OpenRecordset StrSql, rs_sql
'        If rs_sql.EOF Then
'            'Lo creo
'            StrSql = " INSERT INTO titulo (titdesabr,instnro,nivnro) VALUES ("
'            StrSql = StrSql & "'" & Titulo_Descripcion & "'"
'            StrSql = StrSql & "," & Institucion_Nro
'            StrSql = StrSql & "," & Ne_Nro
'            StrSql = StrSql & ")"
'            objConn.Execute StrSql, , adExecuteNoRecords
'
'            Titulo_Nro = getLastIdentity(objConn, "titulo")
'        Else
'            Titulo_Nro = rs_sql!titnro
'        End If
'    Else
'        'Valido que exista el titulo. Si no existe lo creo
'        StrSql = " SELECT * FROM titulo "
'        StrSql = StrSql & " WHERE titdesabr = 'Sin Datos'"
'        StrSql = StrSql & " AND instnro =" & Institucion_Nro
'        StrSql = StrSql & " AND nivnro =" & Ne_Nro
'        If rs_sql.State = adStateOpen Then rs_sql.Close
'        OpenRecordset StrSql, rs_sql
'        If rs_sql.EOF Then
'            'Lo creo
'            StrSql = " INSERT INTO titulo (titdesabr,instnro,nivnro) VALUES ("
'            StrSql = StrSql & "'Sin Datos'"
'            StrSql = StrSql & "," & Institucion_Nro
'            StrSql = StrSql & "," & Ne_Nro
'            StrSql = StrSql & ")"
'            objConn.Execute StrSql, , adExecuteNoRecords
'
'            Titulo_Nro = getLastIdentity(objConn, "titulo")
'        Else
'            Titulo_Nro = rs_sql!titnro
'        End If
'    End If
'
'
'    'Columna 17 - Postgrado. Si / No
'    pos1 = pos2 + 1
'    pos2 = InStr(pos1 + 1, strReg, Separador)
'    tiene_postgrado = IIf(UCase(Mid(strReg, pos1, pos2 - pos1)) = "SI", True, False)
'
'    If tiene_postgrado Then
'        'Columna 18 - Descripcion del Postitulo
'        pos1 = pos2 + 1
'        pos2 = InStr(pos1 + 1, strReg, Separador)
'        PosTitulo_Descripcion = Mid(strReg, pos1, pos2 - pos1)
'        Call AcotarStr(PosTitulo_Descripcion, 40, False)
'
'        If Not IsNull(PosTitulo_Descripcion) Then
'            'Valido que exista el postitulo. Si no existe lo creo
'            StrSql = " SELECT * FROM titulo "
'            StrSql = StrSql & " WHERE titdesabr = '" & PosTitulo_Descripcion & "'"
'            StrSql = StrSql & " AND instnro =" & Institucion_Nro
'            StrSql = StrSql & " AND nivnro =" & Ne_Nro
'            If rs_sql.State = adStateOpen Then rs_sql.Close
'            OpenRecordset StrSql, rs_sql
'            If rs_sql.EOF Then
'                'Lo creo
'                StrSql = " INSERT INTO titulo (titdesabr,instnro,nivnro) VALUES ("
'                StrSql = StrSql & "'" & PosTitulo_Descripcion & "'"
'                StrSql = StrSql & "," & Institucion_Nro
'                StrSql = StrSql & "," & Ne_Nro
'                StrSql = StrSql & ")"
'                objConn.Execute StrSql, , adExecuteNoRecords
'
'                PosTitulo_Nro = getLastIdentity(objConn, "titulo")
'            Else
'                PosTitulo_Nro = rs_sql!titnro
'            End If
'        Else
'            'Valido que exista el titulo. Si no existe lo creo
'            StrSql = " SELECT * FROM titulo "
'            StrSql = StrSql & " WHERE titdesabr = 'Sin Datos'"
'            StrSql = StrSql & " AND instnro =" & Institucion_Nro
'            StrSql = StrSql & " AND nivnro =" & Ne_Nro
'            If rs_sql.State = adStateOpen Then rs_sql.Close
'            OpenRecordset StrSql, rs_sql
'            If rs_sql.EOF Then
'                'Lo creo
'                StrSql = " INSERT INTO titulo (titdesabr,instnro,nivnro) VALUES ("
'                StrSql = StrSql & "'Sin Datos'"
'                StrSql = StrSql & "," & Institucion_Nro
'                StrSql = StrSql & "," & Ne_Nro
'                StrSql = StrSql & ")"
'                objConn.Execute StrSql, , adExecuteNoRecords
'
'                PosTitulo_Nro = getLastIdentity(objConn, "titulo")
'            Else
'                PosTitulo_Nro = rs_sql!titnro
'            End If
'        End If
'    End If
'
'    'Columna 19 - Cargo
'    pos1 = pos2 + 1
'    pos2 = InStr(pos1 + 1, strReg, Separador)
'    Cargo_Descripcion = Mid(strReg, pos1, pos2 - pos1)
'    Call AcotarStr(Cargo_Descripcion, 40, False)
'
'    If Not IsNull(Cargo_Descripcion) Then
'        'Valido que exista el postitulo. Si no existe lo creo
'        StrSql = " SELECT * FROM cargo "
'        StrSql = StrSql & " WHERE cardesabr = '" & Cargo_Descripcion & "'"
'        If rs_sql.State = adStateOpen Then rs_sql.Close
'        OpenRecordset StrSql, rs_sql
'        If rs_sql.EOF Then
'            'Lo creo
'            StrSql = " INSERT INTO cargo (cardesabr) VALUES ("
'            StrSql = StrSql & "'" & Cargo_Descripcion & "'"
'            StrSql = StrSql & ")"
'            objConn.Execute StrSql, , adExecuteNoRecords
'
'            Cargo_Nro = getLastIdentity(objConn, "cargo")
'        Else
'            Cargo_Nro = rs_sql!carnro
'        End If
'    Else
'        'Valido que exista el titulo. Si no existe lo creo
'        StrSql = " SELECT * FROM cargo "
'        StrSql = StrSql & " WHERE titdesabr = 'Sin Datos'"
'        If rs_sql.State = adStateOpen Then rs_sql.Close
'        OpenRecordset StrSql, rs_sql
'        If rs_sql.EOF Then
'            'Lo creo
'            StrSql = " INSERT INTO titulo (cardesabr) VALUES ("
'            StrSql = StrSql & "'Sin Datos'"
'            StrSql = StrSql & ")"
'            objConn.Execute StrSql, , adExecuteNoRecords
'
'            Cargo_Nro = getLastIdentity(objConn, "cargo")
'        Else
'            Cargo_Nro = rs_sql!carnro
'        End If
'    End If
'
'
''''    'Busco la Provincia
''''    StrSql = " SELECT * FROM provincia "
''''    StrSql = StrSql & " WHERE provdesc = '" & Dir_Provincia & "'"
''''    If rs_sql.State = adStateOpen Then rs_sql.Close
''''    OpenRecordset StrSql, rs_sql
''''    If rs_sql.EOF Then
''''        Provincia_Nro = rs_sql!provnro
''''        Pais_Nro = rs_sql!paisnro
''''    Else
''''        Provincia_Nro = 1   'No informada
''''        Pais_Nro = 3        'Argentina
''''    End If
''''
''''    'Busco la localidad
''''    StrSql = " SELECT * FROM localidad "
''''    StrSql = StrSql & " WHERE locdesc = '" & Dir_Localidad & "'"
''''    StrSql = StrSql & " AND provnro = " & Provincia_Nro
''''    If rs_sql.State = adStateOpen Then rs_sql.Close
''''    OpenRecordset StrSql, rs_sql
''''    If rs_sql.EOF Then
''''        Localidad_Nro = rs_sql!locnro
''''    Else
''''        Localidad_Nro = 1   'No Informada
''''    End If
'
''''    Ne_Descripcion = Mid(strReg, pos1, pos2 - pos1)
''''    Call AcotarStr(Ne_Descripcion, 40, False)
''''    If Not IsNull(Ne_Descripcion) Then
''''        'Valido que exista. Si no existe lo creo
''''        StrSql = " SELECT * FROM nivest "
''''        StrSql = StrSql & " WHERE nivdesc = '" & Ne_Descripcion & "'"
''''        If rs_sql.State = adStateOpen Then rs_sql.Close
''''        OpenRecordset StrSql, rs_sql
''''        If rs_sql.EOF Then
''''            'Lo creo
''''            StrSql = " INSERT INTO nivest (nivdesc,nivsist,nivobligatorio,nivestfli) VALUES ("
''''            StrSql = StrSql & "'" & Ne_Descripcion & "'"
''''            StrSql = StrSql & ",0,0,0"
''''            StrSql = StrSql & ")"
''''            objConn.Execute StrSql, , adExecuteNoRecords
''''
''''            Ne_Nro = getLastIdentity(objConn, "nivest")
''''        Else
''''            Ne_Nro = rs_sql!Nivnro
''''        End If
''''    Else
''''        'Valido que exista un registro N\A. Si no existe lo creo
''''        StrSql = " SELECT * FROM nivest "
''''        StrSql = StrSql & " WHERE nivdesc = 'N\A'"
''''        If rs_sql.State = adStateOpen Then rs_sql.Close
''''        OpenRecordset StrSql, rs_sql
''''        If rs_sql.EOF Then
''''            'Lo creo
''''            StrSql = " INSERT INTO nivest (nivdesc,nivsist,nivobligatorio,nivestfli) VALUES ("
''''            StrSql = StrSql & "'N\A'"
''''            StrSql = StrSql & ",0,0,0"
''''            StrSql = StrSql & ")"
''''            objConn.Execute StrSql, , adExecuteNoRecords
''''
''''            Ne_Nro = getLastIdentity(objConn, "nivest")
''''    End If
'
'
'' --------------------------------------------------------------------------------*
'' Inserto
'
'    'Inserto el Tercero
'    StrSql = " INSERT INTO tercero (ternom,terape,terfecnac,tersex,teremail) VALUES ("
'    StrSql = StrSql & "'" & Ter_Nombre & "'"
'    StrSql = StrSql & ",'" & Ter_Apellido & "'"
'    StrSql = StrSql & "," & ConvFecha(Fecha_Nacimiento)
'    StrSql = StrSql & "," & Sexo
'    StrSql = StrSql & ",'" & Email & "'"
'    StrSql = StrSql & ")"
'    objConn.Execute StrSql, , adExecuteNoRecords
'
'    NroTercero = getLastIdentity(objConn, "tercero")
'    Flog.Writeline "Codigo de Tercero = " & NroTercero
'
'    'Inserto en postulantes
'    StrSql = " INSERT INTO pos_postulante(ternro,posfecpres) VALUES("
'    StrSql = StrSql & NroTercero
'    StrSql = StrSql & "," & ConvFecha(Date)
'    StrSql = StrSql & ")"
'    objConn.Execute StrSql, , adExecuteNoRecords
'    Flog.Writeline "Inserte el Postulante - " & Ter_Apellido & " - " & Ter_Nombre
'
'    ' Inserto el Registro correspondiente en ter_tip
'    StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & NroTercero & ",16)"
'    objConn.Execute StrSql, , adExecuteNoRecords
'
'    ' Inserto el Documento
'    If Docu_Nro <> 0 Then
'        StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
'        StrSql = StrSql & " VALUES(" & NroTercero & "," & Docu_Tipo & ",'" & Docu_Nro & "')"
'        objConn.Execute StrSql, , adExecuteNoRecords
'        Flog.Writeline "Inserte el Documento"
'    End If
'
'    ' Inserto el Domicilio
'      StrSql = " INSERT INTO cabdom(tipnro,ternro,domdefault,tidonro) "
'      StrSql = StrSql & " VALUES(1," & NroTercero & ",-1,2)"
'      objConn.Execute StrSql, , adExecuteNoRecords
'
'      NroDom = getLastIdentity(objConn, "cabdom")
'
'      StrSql = " INSERT INTO detdom(domnro,calle,nro,piso,oficdepto,torre,manzana,codigopostal,"
'      StrSql = StrSql & "locnro,provnro,paisnro) "
'      StrSql = StrSql & " VALUES ("
'      StrSql = StrSql & NroDom
'      StrSql = StrSql & ",'" & Dir_Calle & "'"
'      StrSql = StrSql & ",'" & Dir_Nro & "'"
'      StrSql = StrSql & ",'" & Dir_Piso & "'"
'      StrSql = StrSql & ",'" & Dir_OficDepto & "'"
'      StrSql = StrSql & ",'" & Dir_Torre & "'"
'      StrSql = StrSql & ",'" & Dir_Manzana & "'"
'      StrSql = StrSql & ",'" & Dir_CP & "'"
'      StrSql = StrSql & "," & Localidad_Nro
'      StrSql = StrSql & "," & Provincia_Nro
'      StrSql = StrSql & "," & Pais_Nro
'      StrSql = StrSql & ")"
'      objConn.Execute StrSql, , adExecuteNoRecords
'      Flog.Writeline "Inserte el Domicilio "
'
'      If Not IsNull(Tel_Default) Then
'        StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular) "
'        StrSql = StrSql & " VALUES(" & NroDom & ",'" & Tel_Default & "',0,-1,0)"
'        objConn.Execute StrSql, , adExecuteNoRecords
'        Flog.Writeline "Inserte el Telefono Default"
'      End If
'
'      If Not IsNull(Tel_Celular) Then
'        StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular) "
'        StrSql = StrSql & " VALUES(" & NroDom & ",'" & Tel_Celular & "',0,0,-1)"
'        objConn.Execute StrSql, , adExecuteNoRecords
'        Flog.Writeline "Inserte el Telefono Celular "
'      End If
'  End If
'
'    'Inserto los estudios formales
'      StrSql = " INSERT INTO cap_estformal(ternro,nivnro,titnro,instnro,capcomp,capactual)"
'      StrSql = StrSql & " VALUES ("
'      StrSql = StrSql & NroTercero
'      StrSql = StrSql & "," & Ne_Nro
'      StrSql = StrSql & "," & Titulo_Nro
'      StrSql = StrSql & "," & Institucion_Nro
'      StrSql = StrSql & "," & CInt(Ne_Completo)
'      StrSql = StrSql & "," & IIf(tiene_postgrado, 0, -1)
'      StrSql = StrSql & ")"
'      objConn.Execute StrSql, , adExecuteNoRecords
'      Flog.Writeline "Inserte el Nivel de estudio "
'
'    'Postgrados
'    If tiene_postgrado Then
'      StrSql = " INSERT INTO cap_estformal(ternro,nivnro,titnro,instnro,capcomp,capactual)"
'      StrSql = StrSql & " VALUES ("
'      StrSql = StrSql & NroTercero
'      StrSql = StrSql & "," & Ne_Nro
'      StrSql = StrSql & "," & PosTitulo_Nro
'      StrSql = StrSql & "," & Institucion_Nro
'      StrSql = StrSql & "," & CInt(Ne_Completo)
'      StrSql = StrSql & "," & IIf(tiene_postgrado, 0, -1)
'      StrSql = StrSql & ")"
'      objConn.Execute StrSql, , adExecuteNoRecords
'      Flog.Writeline "Inserte el Nivel de estudio - Postgrado"
'    End If
'
'    'Empleos anteriores
'    If Not IsNull(Cargo_Descripcion) Then
'
'    End If
'
'  If rs.State = adStateOpen Then rs.Close
'  If rs_sql.State = adStateOpen Then rs_sql.Close
'
'  Set rs = Nothing
'  Set rs_sql = Nothing
'End Sub


Public Sub LineaModelo_227(ByVal strLinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1            As Integer
Dim pos2            As Integer









End Sub

Public Sub LineaModelo_228(ByVal strLinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    'FGZ -  22/09/2004
    'No USAR. este modelo se usó en el reporte de Declaracion Jurada de la estrella.
    
End Sub

Public Sub LineaModelo_229(ByVal strLinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Interface de Prestamos
' Autor      : FGZ
' Fecha      : 27/09/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1            As Integer
Dim pos2            As Integer

Dim Nro_Legajo As Long
Dim Lin_Prestamo As Long
Dim Descripcion As String
Dim Sucursal As Long
Dim Nro_Comprobante As Long
Dim Cant_Cuotas As Integer
Dim Monto_Total As Single
Dim Anio As Integer
Dim mes As Integer
Dim Fecha_Otorg As Date

Dim Tercero As Long
Dim TipoPrestamo As Long
Dim lnprenro As Long
Dim Moneda As Long
Dim PliqNro As Long
Dim Encontro As Boolean
Dim EstNro As Long

Dim rs_Empleado As New ADODB.Recordset
Dim rs_TipoPrestamo As New ADODB.Recordset
Dim rs_Sucursal As New ADODB.Recordset
Dim rs_Monedas As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Estado As New ADODB.Recordset
Dim rs_pre_linea  As New ADODB.Recordset

' El formato es:
'   Legajo; Linea de Prestamo; Descripcion; Sucursal; Nro Comprobante; Cant. de Cuotas;
    'Monto Total; Año; Mes; Fecha Otorg.

    'Nro de Legajo
    pos1 = 1
    pos2 = InStr(pos1, strLinea, Separador)
    If IsNumeric(Mid$(strLinea, pos1, pos2 - pos1)) Then
        Nro_Legajo = Mid$(strLinea, pos1, pos2 - pos1)
    Else
        InsertaError 1, 8
        Exit Sub
    End If
    
    'Linea de Prestamo
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    Lin_Prestamo = Mid(strLinea, pos1, pos2 - pos1)

    'Descripcion
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    Descripcion = Mid(strLinea, pos1, pos2 - pos1)

    'Sucursal
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    Sucursal = Mid(strLinea, pos1, pos2 - pos1)

    'Nro Comprobante
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    Nro_Comprobante = Mid(strLinea, pos1, pos2 - pos1)

    'Cantidad de Cuotas
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    Cant_Cuotas = Mid(strLinea, pos1, pos2 - pos1)

    'Monto Total
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    Monto_Total = Mid(strLinea, pos1, pos2 - pos1)
    Monto_Total = CSng(Replace(CStr(Monto_Total), SeparadorDecimal, "."))
    
    'Año
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    Anio = Mid(strLinea, pos1, pos2 - pos1)

    'Mes
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    mes = Mid(strLinea, pos1, pos2 - pos1)

    'Fecha de Otorgamiento
    pos1 = pos2 + 1
    pos2 = Len(strLinea)
    Fecha_Otorg = Mid(strLinea, pos1, pos2)

' ----------------------------------
'Validaciones

'Que exista el legajo
StrSql = "SELECT * FROM empleado where empleg = " & Nro_Legajo
OpenRecordset StrSql, rs_Empleado
If rs_Empleado.EOF Then
    Flog.Writeline "No se encontro el legajo " & Nro_Legajo
    InsertaError 1, 8
    Exit Sub
Else
    Tercero = rs_Empleado!ternro
End If

'que exista la sucursal
StrSql = "SELECT * FROM sucursal where estrnro = " & Sucursal
OpenRecordset StrSql, rs_Sucursal
If rs_Sucursal.EOF Then
    Flog.Writeline "No se encontro la sucursal " & Sucursal
    InsertaError 4, 56
    Exit Sub
End If

'Que el monto sea numerico
If Not IsNumeric(Monto_Total) Then
    Flog.Writeline "El monto no es numerico " & Monto_Total
    InsertaError 7, 5
    Exit Sub
End If

'que el mes sea valido
If IsNumeric(mes) Then
    If mes > 12 Or mes < 1 Then
        Flog.Writeline "El mes es incorrecto " & mes
        InsertaError 9, 42
        Exit Sub
    End If
Else
    Flog.Writeline "El mes es incorrecto " & mes
    InsertaError 9, 42
    Exit Sub
End If

'que el año sea valido
If Not IsNumeric(Anio) Then
    Flog.Writeline "El año es incorrecto " & Anio
    InsertaError 8, 3
    Exit Sub
End If

'que la fecha de otorgamiento sea una fecha
If Not IsDate(Fecha_Otorg) Then
    Flog.Writeline "La fecha es incorrecta " & Fecha_Otorg
    InsertaError 10, 4
    Exit Sub
End If

'Busco el Estado de prestamos (Nuevo)
StrSql = "SELECT * FROM estadopre ORDER BY estnro "
OpenRecordset StrSql, rs_Estado
If rs_Estado.EOF Then
    Flog.Writeline "No se encontro el estado Pendiente para los prestamos "
    EstNro = 0
Else
    EstNro = rs_Estado!EstNro
End If

'Busco la primer moneda
StrSql = "SELECT * FROM moneda"
OpenRecordset StrSql, rs_Monedas
If rs_Monedas.EOF Then
    Flog.Writeline "No se encontro ninguna Moneda"
    'InsertaError 4, 56
    'Exit Sub
    Moneda = 0
Else
    Moneda = rs_Monedas!monnro
End If

'Busco el periodo de liquidacion al que le corresponde
StrSql = "SELECT * FROM periodo"
StrSql = StrSql & " WHERE pliqmes = " & mes
StrSql = StrSql & " AND pliqanio = " & Anio
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.Writeline "No se encontro periodo de liquidacion asociado"
    'InsertaError 4, 56
    'Exit Sub
    PliqNro = 0
Else
    PliqNro = rs_Periodo!PliqNro
End If


'Busco la linea de prestamo
StrSql = "SELECT * FROM pre_linea WHERE lnprenro =" & Lin_Prestamo
OpenRecordset StrSql, rs_pre_linea
If rs_pre_linea.EOF Then
    Flog.Writeline "No se encontro la linea de Prestamo"
    InsertaError 2, 96
    Exit Sub
Else
    lnprenro = rs_pre_linea!lnprenro
End If

' ---------------------------------------------
'Inserto

'inserto en prestamo
StrSql = "INSERT INTO prestamo ("
StrSql = StrSql & "predesc,preimp,precantcuo,ternro,quincenal,premes,preanio,monnro"
StrSql = StrSql & ",lnprenro,iduser,prefecotor,sucursal,pliqnro,precompr"
StrSql = StrSql & ",pretna,preiva,preotrosgas,prediavto "
StrSql = StrSql & ") VALUES ("
StrSql = StrSql & "'" & Left(Descripcion, 60) & "'"
StrSql = StrSql & "," & Monto_Total
StrSql = StrSql & "," & Cant_Cuotas
StrSql = StrSql & "," & Tercero
StrSql = StrSql & ",0"
StrSql = StrSql & "," & mes
StrSql = StrSql & "," & Anio
StrSql = StrSql & "," & Moneda

StrSql = StrSql & "," & lnprenro
StrSql = StrSql & ",0"
StrSql = StrSql & "," & ConvFecha(Fecha_Otorg)
StrSql = StrSql & "," & Sucursal
StrSql = StrSql & "," & PliqNro
StrSql = StrSql & ",'" & Nro_Comprobante & "'"

StrSql = StrSql & ",0,0,0,1 "

StrSql = StrSql & " )"
objConn.Execute StrSql, , adExecuteNoRecords


'Cierro y libero
If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
If rs_TipoPrestamo.State = adStateOpen Then rs_TipoPrestamo.Close
If rs_Sucursal.State = adStateOpen Then rs_Sucursal.Close
If rs_Monedas.State = adStateOpen Then rs_Monedas.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Estado.State = adStateOpen Then rs_Estado.Close
If rs_pre_linea.State = adStateOpen Then rs_pre_linea.Close

Set rs_Empleado = Nothing
Set rs_TipoPrestamo = Nothing
Set rs_Sucursal = Nothing
Set rs_Monedas = Nothing
Set rs_Periodo = Nothing
Set rs_Estado = Nothing
Set rs_pre_linea = Nothing

End Sub


Public Sub LineaModelo_230(ByVal strLinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Interface de Dias Pedidos de Vacaciones
' Autor      : FGZ
' Fecha      : 04/10/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
' El formato es:
'   Legajo; Nombre y Apellido; Antiguedad Años; Antiguedad Meses; Dias Pendientes; Dias Correspondientes;
'   Total Dias; Fecha Desde; Fecha Hasta
' ---------------------------------------------------------------------------------------------
Dim pos1            As Integer
Dim pos2            As Integer

Dim Nro_Legajo As Long
Dim Nombre_Apellido As String
Dim Ant_Anios As Integer
Dim Ant_Meses As Integer
Dim Aux_Dias_Pendientes As Integer
Dim Aux_Dias_Correspondientes As Integer
Dim Aux_TipoVacacion As String
Dim Total_Dias As Integer
Dim Fecha_Desde As Date
Dim Fecha_Hasta As Date
Dim Tercero As Long

Dim Aux_Fecha_Desde As Date

Dim rs_tipovacac As New ADODB.Recordset
Dim rs_vacdiascor As New ADODB.Recordset
Dim rsDias As New ADODB.Recordset
Dim rsVac As New ADODB.Recordset

Dim diascoract As Integer
Dim diastom As Integer
Dim diascorant As Integer
Dim diasdebe As Integer
Dim diastot As Integer
Dim diasyaped As Integer
Dim diaspend As Integer

Dim nroTipvac As Long
Dim Hasta As Date
Dim totferiados As Integer
Dim tothabiles As Integer
Dim totNohabiles As Integer
Dim NroVac As Long

Dim rs_Empleado As New ADODB.Recordset
Dim rs_Periodos_Vac As New ADODB.Recordset


'Actio el Manejador de Errores Local
On Error GoTo Manejador_Local

    'Nro de Legajo
    pos1 = 1
    pos2 = InStr(pos1, strLinea, Separador)
    If IsNumeric(Mid$(strLinea, pos1, pos2 - pos1)) Then
        Nro_Legajo = Mid$(strLinea, pos1, pos2 - pos1)
    Else
        InsertaError 1, 8
        Exit Sub
    End If
    
    'Nombre y Apellido
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    Nombre_Apellido = Mid(strLinea, pos1, pos2 - pos1)

    'Antiguedad Años
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    Ant_Anios = Mid(strLinea, pos1, pos2 - pos1)

    'Antiguedad Mese
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    Ant_Meses = Mid(strLinea, pos1, pos2 - pos1)

    'Tipo de Vacaciones
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    Aux_TipoVacacion = Mid(strLinea, pos1, pos2 - pos1)

    'Dias_Pendientes
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    Aux_Dias_Pendientes = Mid(strLinea, pos1, pos2 - pos1)

    'Dias Correspondientes
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    Aux_Dias_Correspondientes = Mid(strLinea, pos1, pos2 - pos1)

    'Total de Dias
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    Total_Dias = Mid(strLinea, pos1, pos2 - pos1)
    
    'Fecha Desde
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    Fecha_Desde = Mid(strLinea, pos1, pos2 - pos1)
    
    'Fecha Hasta
    pos1 = pos2 + 1
    pos2 = Len(strLinea)
    Fecha_Hasta = Mid(strLinea, pos1, pos2)

    ' ----------------------------------
    'Validaciones
    
    'Que exista el legajo
    StrSql = "SELECT * FROM empleado where empleg = " & Nro_Legajo
    OpenRecordset StrSql, rs_Empleado
    If rs_Empleado.EOF Then
        Flog.Writeline "No se encontro el legajo " & Nro_Legajo
        InsertaError 1, 8
        Exit Sub
    Else
        Tercero = rs_Empleado!ternro
    End If
    
    Aux_Fecha_Desde = Fecha_Desde
    
    'Busco todos los Periodos Involucrados entre las Fechas
    StrSql = "SELECT * FROM vacacion "
    StrSql = StrSql & " WHERE vacfecdesde <= " & ConvFecha(Fecha_Hasta)
    StrSql = StrSql & " AND  vacfechasta >= " & ConvFecha(Fecha_Desde)
    StrSql = StrSql & " ORDER BY vacnro"
    OpenRecordset StrSql, rs_Periodos_Vac
    
    Do While Not rs_Periodos_Vac.EOF And Aux_Fecha_Desde < Fecha_Hasta
        'si le fecha en la que se va a generar los dias pedidos estan fuera del rengo de fechas del periodo
        ' no se procesan
        If Aux_Fecha_Desde >= rs_Periodos_Vac!vacfecdesde And Aux_Fecha_Desde <= rs_Periodos_Vac!vacfechasta Then
            diascoract = 0
            diastom = 0
            diascorant = 0
            diasdebe = 0
            diastot = 0
            diasyaped = 0
            diaspend = 0
            
            Flog.Writeline "Periodo de Vacaciones:" & rs_Periodos_Vac!vacnro & " " & rs_Periodos_Vac!vacdesc
            
            NroVac = rs_Periodos_Vac!vacnro
            
            StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Tercero & " AND vacnro = " & NroVac
            OpenRecordset StrSql, rs_vacdiascor
            If Not rs_vacdiascor.EOF Then
     
                StrSql = "SELECT * FROM tipovacac WHERE tipvacnro = " & rs_vacdiascor!tipvacnro
                OpenRecordset StrSql, rs_tipovacac
                If Not rs_tipovacac.EOF Then
                    nroTipvac = rs_tipovacac!tipvacnro
                End If
                
                diascoract = rs_vacdiascor!vdiascorcant ' dias corresp al periodo actual
                
                StrSql = "SELECT * FROM vacacion WHERE vacnro <> " & NroVac & " and vacfechasta < " & ConvFecha(Fecha_Desde)
                OpenRecordset StrSql, rsVac
                Do While Not rsVac.EOF
                
                    diastom = 0
                    
                    StrSql = "SELECT * FROM lic_vacacion " & _
                             " INNER JOIN emp_lic ON emp_lic.emp_licnro = lic_vacacion.emp_licnro " & _
                             " WHERE lic_vacacion.vacnro = " & rsVac!vacnro & " AND emp_lic.empleado = " & Tercero
                    OpenRecordset StrSql, rsDias
                    Do While Not rsDias.EOF
                        diastom = diastom + rsDias!elcantdias
                        rsDias.MoveNext
                    Loop
                    diascorant = 0
                    StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Tercero & " AND vacnro = " & rsVac!vacnro
                    OpenRecordset StrSql, rs
                    If Not rs.EOF Then diascorant = rs!vdiascorcant
                    diasdebe = diasdebe + (diascorant - diastom)
                    
                    rsVac.MoveNext
                Loop
                
                diastot = diascoract + diasdebe
            End If
            
            
            ' Busco los pedidos de ese periodo
            StrSql = "SELECT * FROM vacdiasped WHERE ternro = " & Tercero & " AND vacnro = " & NroVac
            OpenRecordset StrSql, objRs
            Do While Not objRs.EOF
                'diasyaped = diasyaped + objRs!vdiapedcant
                diasyaped = diasyaped + objRs!vdiaspedhabiles
                objRs.MoveNext
            Loop
            
            diaspend = diastot - diasyaped
            If diaspend > 0 Then
            
                Call DiasPedidos(NroVac, nroTipvac, Aux_Fecha_Desde, Fecha_Hasta, Hasta, Tercero, diaspend, tothabiles, totNohabiles, totferiados)
                
                StrSql = "INSERT INTO vacdiasped (vdiapedhasta,vdiaspedferiados,vdiaspedhabiles,vdiaspednohabiles,ternro,vacnro,vdiapedcant,vdiapeddesde,vdiaspedestado) VALUES (" & _
                          ConvFecha(Hasta) & "," & totferiados & "," & tothabiles & "," & totNohabiles & "," & Tercero & "," & NroVac & "," & (totferiados + tothabiles + totNohabiles) & "," & ConvFecha(Aux_Fecha_Desde) & ",-1)"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Aux_Fecha_Desde = Hasta + 1
            End If
        Else
            Flog.Writeline "La fecha en la que se va a generar los dias pedidos estan fuera del rengo de fechas del periodo " & Aux_Fecha_Desde
        End If
            
        rs_Periodos_Vac.MoveNext
    Loop
    
    If Aux_Fecha_Desde < Fecha_Hasta Then
        Flog.Writeline "Quedaron " & DateDiff("d", Fecha_Hasta, Aux_Fecha_Desde) + 1 & " días sin generar "
    End If

Exit Sub

'Desactivo el manejador de Errores Local
On Error GoTo 0

Manejador_Local:
    Flog.Writeline "Error de Formato "
    InsertaError 1, 7
    Exit Sub

End Sub


Private Sub DiasPedidos(ByVal NroVac As Long, ByVal TipoVac As Long, ByVal FechaInicial As Date, ByVal FechaFinal As Date, ByRef Fecha As Date, ByVal ternro As Long, ByRef cant As Integer, ByRef cHabiles As Integer, ByRef cNoHabiles As Integer, ByRef cFeriados As Integer)
'calcula la cantidad de dias y la fecha hasta correspondiente al periodo
'de acuerdo al tipo de vacacion

Dim i As Integer
Dim j As Integer
Dim objFeriado As New Feriado
Dim DHabiles(1 To 7) As Boolean
Dim esFeriado As Boolean
Dim objRs As New ADODB.Recordset
Dim ExcluyeFeriados As Boolean

    StrSql = "SELECT * FROM tipovacac WHERE tipvacnro = " & TipoVac
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        DHabiles(1) = objRs!tpvhabiles__1
        DHabiles(2) = objRs!tpvhabiles__2
        DHabiles(3) = objRs!tpvhabiles__3
        DHabiles(4) = objRs!tpvhabiles__4
        DHabiles(5) = objRs!tpvhabiles__5
        DHabiles(6) = objRs!tpvhabiles__6
        DHabiles(7) = objRs!tpvhabiles__7
    
        ExcluyeFeriados = CBool(objRs!tpvferiado)
    Else
        Flog.Writeline "No se encontro el tipo de Vacacion " & TipoVac
        Exit Sub
    End If
    
    
    Set objFeriado.Conexion = objConn
    Set objFeriado.ConexionTraza = objConn
    
    i = 0
    j = 0
    cHabiles = 0
    cNoHabiles = 0
    cFeriados = 0
    
    Fecha = FechaInicial
    
    Do While Fecha <= FechaFinal
        esFeriado = objFeriado.Feriado(Fecha, ternro, False)
        
        If (esFeriado) And Not ExcluyeFeriados Then
            cFeriados = cFeriados + 1
        Else
            If DHabiles(Weekday(Fecha)) Or (esFeriado And ExcluyeFeriados) Then
                i = i + 1
                If DHabiles(Weekday(Fecha)) Then
                    cHabiles = cHabiles + 1
                End If
            Else
                cNoHabiles = cNoHabiles + 1
            End If
        End If
        Fecha = DateAdd("d", 1, Fecha)
    Loop
    Fecha = DateAdd("d", -1, Fecha)
    
    cant = cHabiles '+ cNoHabiles + cFeriados
    Set objFeriado = Nothing
End Sub


Public Sub LineaModelo_231(ByVal strLinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Interface de Banco Nacion. Levanta y crea las cuentas para los empleados.
' Autor      : FGZ
' Fecha      : 14/10/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Registro As TR_Datos_Bancarios
Dim CBU As TR_CBU
Dim CBU_Bloque1 As TR_CBU_Bloque_1
Dim CBU_Bloque2 As TR_CBU_Bloque_2

Dim pos1 As Integer
Dim pos2 As Integer
    
Dim Tercero As Long
Dim FechaDesde As String
Dim FechaHasta As String
Dim Tipo_Cuenta As Long
Dim Aux_Tipo_Cuenta As Long
Dim Nro_Reporte As Long

Dim rs_Empleado As New ADODB.Recordset
Dim rs_CtaBancaria As New ADODB.Recordset
Dim rs_FormaPago As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset

' El formato es: segun formato del tipo de datos TR_Datos_Bancarios
        
    'Proceso        long 2
    pos1 = 1
    pos2 = 2 + 1
    Registro.Proceso = Mid$(strLinea, pos1, pos2 - pos1)
    
    'Servicio       long 4
    pos1 = 3
    pos2 = 6 + 1
    Registro.Servicio = Mid$(strLinea, pos1, pos2 - pos1)
    
    'Sucursal       long 4
    pos1 = 7
    pos2 = 10 + 1
    Registro.Sucursal = Mid$(strLinea, pos1, pos2 - pos1)
    
    'Legajo         long 20
    pos1 = 11
    pos2 = 30 + 1
    Registro.Legajo = Mid$(strLinea, pos1, pos2 - pos1)
    
    'Moneda         long 1
    pos1 = 31
    pos2 = 31 + 1
    Registro.Moneda = Mid$(strLinea, pos1, pos2 - pos1)
    
    'Titularidad    long 2
    pos1 = 32
    pos2 = 33 + 1
    Registro.Titularidad = Mid$(strLinea, pos1, pos2 - pos1)
    
    'CBU long 22 -
        'Bloque 1
        '       Codigo entidad:  "011"  - long 3
                pos1 = 34
                pos2 = 36 + 1
                CBU_Bloque1.Codigo_Entidad = Mid$(strLinea, pos1, pos2 - pos1)
        '       Codigo Sucursal: "BCRA" - long 4
                pos1 = 37
                pos2 = 40 + 1
                CBU_Bloque1.Codigo_Sucursal = Mid$(strLinea, pos1, pos2 - pos1)
        '       Digito Verif. Bloque 1  - long 1
                pos1 = 41
                pos2 = 41 + 1
                CBU_Bloque1.Digito_Verificador = Mid$(strLinea, pos1, pos2 - pos1)
    
        'Bloque 2
        '       Tipo de Cuenta:         - long 1  (2 = CC, 3 = CA y 4 = CCE)
                pos1 = 42
                pos2 = 42 + 1
                CBU_Bloque2.Cuenta_Tipo = Mid$(strLinea, pos1, pos2 - pos1)
        '       Moneda de la cuenta:    - long 1  (0 = Pesos, 1 = Dolares y 3 = Lecop)
                pos1 = 43
                pos2 = 43 + 1
                CBU_Bloque2.Moneda = Mid$(strLinea, pos1, pos2 - pos1)
        '       Nro de la cuenta        - long 11
                pos1 = 44
                pos2 = 54 + 1
                CBU_Bloque2.Cuenta_Nro = Mid$(strLinea, pos1, pos2 - pos1)
        '       Digito Verif. Bloque 2  - long 1
                pos1 = 55
                pos2 = 55 + 1
                CBU_Bloque2.Digito_Verificador = Mid$(strLinea, pos1, pos2 - pos1)
    CBU.Bloque1 = CBU_Bloque1.Codigo_Entidad & CBU_Bloque1.Codigo_Sucursal & CBU_Bloque1.Digito_Verificador
    CBU.Bloque2 = CBU_Bloque2.Cuenta_Tipo & CBU_Bloque2.Moneda & CBU_Bloque2.Cuenta_Nro & CBU_Bloque2.Digito_Verificador
    Registro.CBU = CBU.Bloque1 & CBU.Bloque2
    
    'Cuenta_Electronica     long 19 - (nro de tarjeta de debito)
    pos1 = 56
    pos2 = 74 + 1
    Registro.Cuenta_Electronica = Mid$(strLinea, pos1, pos2 - pos1)
    'Tarjeta_1er_Titular    long 19 -
    pos1 = 75
    pos2 = 93 + 1
    Registro.Tarjeta_1er_Titular = Mid$(strLinea, pos1, pos2 - pos1)
    'Tarjeta_2do_Titular    long 19 -
    pos1 = 94
    pos2 = 112 + 1
    Registro.Tarjeta_2do_Titular = Mid$(strLinea, pos1, pos2 - pos1)
    'Doc_Tipo As String     long 2  -
    pos1 = 113
    pos2 = 114 + 1
    Registro.Doc_Tipo = Mid$(strLinea, pos1, pos2 - pos1)
    'Doc_Nro                long 11 -
    pos1 = 115
    pos2 = 125 + 1
    Registro.Doc_Nro = Mid$(strLinea, pos1, pos2 - pos1)
    'Filler                 long 5  -
    pos1 = 126
    pos2 = 130 + 1
    Registro.Filler = Mid$(strLinea, pos1, pos2 - pos1)
    
' ====================================================================
'   Validar los parametros Levantados

'Que exista el legajo
StrSql = "SELECT * FROM empleado where empleg = " & Registro.Legajo
OpenRecordset StrSql, rs_Empleado
If rs_Empleado.EOF Then
    Flog.Writeline Espacios(Tabulador * 2) & "No se encontro el legajo " & Registro.Legajo
    InsertaError 1, 8
    Exit Sub
Else
    Tercero = rs_Empleado!ternro
End If

'Configuracion del Reporte
If Primera_Vez Then
    Nro_Reporte = 112
    StrSql = "SELECT * FROM confrep WHERE confnrocol = 3 AND repnro = " & Nro_Reporte
    OpenRecordset StrSql, rs_Confrep
    If rs_Confrep.EOF Then
        Flog.Writeline Espacios(Tabulador * 2) & "CONFREP " & Nro_Reporte & " No se encontró la configuración del Banco"
        InsertaError 0, 60
        Exit Sub
    Else
        Banco = rs_Confrep!confval
        Primera_Vez = False
    End If
End If

'Busco el tipo de cuenta
Select Case CBU_Bloque2.Cuenta_Tipo
Case 2:
    Aux_Tipo_Cuenta = 4
Case 3:
    Aux_Tipo_Cuenta = 2
Case 4:
    Aux_Tipo_Cuenta = 4
Case Else
    Aux_Tipo_Cuenta = 4
End Select

StrSql = "SELECT * FROM formapago WHERE fpagbanc = -1 "
StrSql = StrSql & " AND fpagnro = " & Aux_Tipo_Cuenta
If rs_FormaPago.State = adStateOpen Then rs_FormaPago.Close
OpenRecordset StrSql, rs_FormaPago
If rs_FormaPago.EOF Then
    Flog.Writeline Espacios(Tabulador * 2) & "No se encontro el tipo de cuenta " & CBU_Bloque2.Cuenta_Tipo
    InsertaError 7, 8
    Exit Sub
Else
    Tipo_Cuenta = rs_FormaPago!fpagnro
End If


'=============================================================
'Busco si existe una cuenta para ese legajo para el mismo banco del mismo tipo de cuenta y activa
StrSql = "SELECT * FROM ctabancaria"
StrSql = StrSql & " WHERE ctabancaria.ternro =" & Tercero
StrSql = StrSql & " AND ctabestado = -1 "
StrSql = StrSql & " AND banco =" & Banco
StrSql = StrSql & " AND fpagnro =" & Tipo_Cuenta
If rs_CtaBancaria.State = adStateOpen Then rs_CtaBancaria.Close
OpenRecordset StrSql, rs_CtaBancaria
If Not rs_CtaBancaria.EOF Then
    Flog.Writeline Espacios(Tabulador * 2) & "No se puede crear la cuenta. Ya existe un cuenta del mismo tipo para el mismo banco activa"
    InsertaError 1, 97
Else
    StrSql = "INSERT INTO ctabancaria (" & _
             " ternro,fpagnro,banco,ctabestado, ctabnro, ctabcbu, ctabporc" & _
             ") VALUES (" & Tercero & _
             "," & Tipo_Cuenta & _
             "," & Banco & _
             ",-1" & _
             ",'" & CBU_Bloque2.Cuenta_Nro & "'" & _
             ",'" & Registro.CBU & "'" & _
             ",100" & _
             " )"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline Espacios(Tabulador * 2) & "Cuenta Creada"
End If
    
'Cierro todo y libero
If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
If rs_CtaBancaria.State = adStateOpen Then rs_CtaBancaria.Close
If rs_FormaPago.State = adStateOpen Then rs_FormaPago.Close
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close

Set rs_CtaBancaria = Nothing
Set rs_Empleado = Nothing
Set rs_FormaPago = Nothing
Set rs_Confrep = Nothing
End Sub


Private Sub ValidarLocalidad(Localidad As String, ByRef nro_localidad As Integer, nro_pais As Integer, nro_provincia As Integer)
Dim rs_sub As New ADODB.Recordset
Dim Sql_Ins As String
Dim SQL_Val As String

If Not EsNulo(Localidad) Then
    StrSql = " SELECT * FROM localidad WHERE locdesc = '" & Localidad & "'"
'    If nro_pais <> 0 Then
'        StrSql = StrSql & " AND paisnro = " & nro_pais
'    End If
'
'    If nro_provincia <> 0 Then
'        StrSql = StrSql & " AND provnro = " & nro_provincia
'    End If
    OpenRecordset StrSql, rs_sub
    
    If rs_sub.EOF Then
    
        Sql_Ins = " INSERT INTO localidad(locdesc"
        SQL_Val = " VALUES('" & Localidad & "'"
    
        If nro_pais <> 0 Then
        
            Sql_Ins = Sql_Ins & ",paisnro"
            SQL_Val = SQL_Val & "," & nro_pais
        
        End If
    
        If nro_provincia <> 0 Then
            Sql_Ins = Sql_Ins & ",provnro"
            SQL_Val = SQL_Val & "," & nro_provincia
        End If
        
        StrSql = Sql_Ins & ")" & SQL_Val & ")"
        
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = " SELECT MAX(locnro) AS MaxLoc FROM localidad "
        'StrSql = " SELECT @@IDENTITY AS MaxLoc "
        OpenRecordset StrSql, rs_sub
            
        nro_localidad = rs_sub!MaxLoc
        
    Else
    
        nro_localidad = rs_sub!locnro
    
    End If
End If
End Sub

Private Sub ValidarPartido(Partido As String, ByRef nro_partido As Integer)

Dim rs_sub As New ADODB.Recordset

If Not EsNulo(Partido) Then
    StrSql = " SELECT * FROM partido WHERE partnom = '" & Partido & "'"
    OpenRecordset StrSql, rs_sub
    
    If rs_sub.EOF Then
    
        StrSql = "INSERT INTO partido(partnom) VALUES('"
        StrSql = StrSql & Partido & "')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = " SELECT MAX(partnro) AS MaxPart FROM partido "
        'StrSql = " SELECT @@IDENTITY AS MaxPart "
        OpenRecordset StrSql, rs_sub
            
        nro_partido = rs_sub!MaxPart
    
    Else
        
        nro_partido = rs_sub!partnro
    
    End If
End If
End Sub

Private Sub ValidarZona(Zona As String, ByRef nro_zona As Integer, nro_provincia As Integer)

Dim rs_sub As New ADODB.Recordset

    If Not EsNulo(Zona) Then
        StrSql = " SELECT * FROM zona WHERE zonadesc = '" & Zona & "' AND provnro = " & nro_provincia
        OpenRecordset StrSql, rs_sub
        
        If rs_sub.EOF Then
        
            StrSql = "INSERT INTO zona(zonadesc,provnro) VALUES('"
            StrSql = StrSql & Zona & "'," & nro_provincia & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
'            StrSql = " SELECT MAX(zona) AS MaxZona FROM zona "
'            'StrSql = " SELECT @@IDENTITY AS MaxZona "
'            OpenRecordset StrSql, rs_sub
'
'            nro_zona = rs_sub!MaxZona
            nro_zona = getLastIdentity(objConn, "zona")
        Else
            
            nro_zona = rs_sub!zonanro
        
        End If
    End If

End Sub

Private Sub ValidarProvincia(Provincia As String, ByRef nro_provincia As Integer, nro_pais As Integer)

Dim rs_sub As New ADODB.Recordset

If Not EsNulo(Provincia) Then
    'StrSql = " SELECT * FROM provincia WHERE provdesc = '" & Provincia & "' AND paisnro = " & nro_pais
    StrSql = " SELECT * FROM provincia WHERE provdesc = '" & Provincia & "'"
    OpenRecordset StrSql, rs_sub
    
    If rs_sub.EOF Then
    
        StrSql = "INSERT INTO provincia(provdesc,paisnro) VALUES('"
        StrSql = StrSql & Provincia & "'," & nro_pais & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = " SELECT MAX(provnro) AS MaxProv FROM provincia "
        'StrSql = " SELECT @@IDENTITY AS MaxProv "
        OpenRecordset StrSql, rs_sub
            
        nro_provincia = rs_sub!MaxProv
    
    Else
        
        nro_provincia = rs_sub!provnro
    
    End If
End If
End Sub

Private Sub ValidarPais(Pais As String, ByRef nro_pais As Integer)

Dim rs_sub As New ADODB.Recordset

    If Not EsNulo(Pais) Then
        StrSql = " SELECT * FROM pais WHERE paisdesc = '" & Pais & "'"
        OpenRecordset StrSql, rs_sub
        
        If rs_sub.EOF Then
        
            StrSql = "INSERT INTO pais(paisdesc,paisdef) VALUES('"
            StrSql = StrSql & Pais & "',0)"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(paisnro) AS MaxPais FROM pais "
            'StrSql = " SELECT @@IDENTITY AS MaxPais "
            OpenRecordset StrSql, rs_sub
                
            nro_pais = rs_sub!MaxPais
        
        Else
            
            nro_pais = rs_sub!paisnro
        
        End If
    End If


End Sub

Private Sub ValidaEstructura(TipoEstr As Integer, ByRef valor As String, ByRef CodEst As Integer, ByRef Inserto_estr As Boolean)

Dim Rs_Estr As New ADODB.Recordset

Dim d_estructura As String
Dim CodExt As String
Dim l_pos1 As Integer
Dim l_pos2 As Integer


    If InStr(1, valor, "$") > 0 Then
        l_pos1 = InStr(1, valor, "$")
        l_pos2 = Len(valor)
    
        d_estructura = Mid(valor, l_pos1 + 2, l_pos2)
        If l_pos1 <> 0 Then
            CodExt = Mid(valor, 1, l_pos1 - 1)
        Else
            CodExt = ""
        End If
    Else
        d_estructura = valor
        CodExt = ""
    End If
    
    valor = d_estructura
    
    StrSql = " SELECT estrnro FROM estructura WHERE estructura.estrdabr = '" & d_estructura & "'"
    StrSql = StrSql & " AND estructura.tenro = " & TipoEstr
    OpenRecordset StrSql, Rs_Estr
        
    If Not Rs_Estr.EOF Then
                
            CodEst = Rs_Estr!estrnro
            Inserto_estr = False
            
    Else
            StrSql = " INSERT INTO estructura(tenro,estrdabr,empnro,estrest,estrcodext)"
            StrSql = StrSql & " VALUES(" & TipoEstr & ",'" & d_estructura & "',1,-1,'" & CodExt & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(estrnro) AS MaxEst FROM estructura "
            'StrSql = " SELECT @@IDENTITY AS MaxEst "
            OpenRecordset StrSql, Rs_Estr
            
            CodEst = Rs_Estr!MaxEst
            Inserto_estr = True
    End If


End Sub
Private Sub CreaComplemento(TipoEstr As Integer, CodTer As Integer, CodEstr As Integer, valor As String)


  Select Case TipoEstr

    Case 1
        Complementos1 CodTer, CodEstr
    Case 3
        Complementos3 CodTer, CodEstr
    Case 4
        Complementos4 CodEstr, valor
    Case 10
        Complementos10 CodTer, CodEstr, valor
    Case 15
        Complementos15 CodTer, CodEstr
    Case 16
        Complementos16 CodTer, CodEstr
    Case 17
        Complementos17 CodTer, CodEstr, valor
    Case 18
        Complementos18 CodTer, CodEstr, valor
    Case 19
        Complementos19 CodEstr
    Case 22
        Complementos22 CodTer, CodEstr, valor
    Case 23
        Complementos23 CodTer, CodEstr, valor
    Case 40
        Complementos40 CodTer, CodEstr, valor
    Case 41
        Complementos41 CodTer, CodEstr, valor

  End Select
 
End Sub
Private Sub Complementos1(CodTer As Integer, CodEstr As Integer)

    StrSql = " INSERT INTO sucursal(estrnro,ternro,sucest) VALUES(" & CodEstr & "," & CodTer & ",-1)"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Private Sub Complementos3(CodTer As Integer, CodEstr As Integer)

    StrSql = " INSERT INTO categoria(estrnro,convnro) VALUES(" & CodEstr & "," & CodTer & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Private Sub Complementos4(CodEstr As Integer, valor As String)

    StrSql = " INSERT INTO puesto(estrnro,puedesc,puenroreemp) VALUES(" & CodEstr & ",'" & valor & "',0)"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Private Sub Complementos10(CodTer As Integer, CodEstr As Integer, valor As String)

    StrSql = " INSERT INTO empresa(estrnro,ternro,empnom) VALUES(" & CodEstr & "," & CodTer & ",'" & valor & "')"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Private Sub Complementos15(CodTer As Integer, CodEstr As Integer)

    ' Hay que crear un Tipo de Caja de Jubilacion "Migracion"

    StrSql = " INSERT INTO cajjub(estrnro,ternro,cajest,ticnro) VALUES(" & CodEstr & "," & CodTer & "-1,1)"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Private Sub Complementos16(CodTer As Integer, CodEstr As Integer)

    StrSql = " INSERT INTO gremio(estrnro,ternro) VALUES(" & CodEstr & "," & CodTer & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Private Sub Complementos17(CodTer As Integer, CodEstr As Integer, valor As String)

    StrSql = " INSERT INTO osocial(ternro,osdesc) VALUES(" & CodTer & ",'" & valor & "')"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    StrSql = " INSERT INTO replica_estr(origen,estrnro) VALUES (" & CodTer & "," & CodEstr & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    

End Sub

Private Sub Complementos18(CodTer As Integer, CodEstr As Integer, valor As String)
Dim rs_tipocont As New ADODB.Recordset
Dim rs_TC As New ADODB.Recordset

    
    StrSql = "SELECT * FROM tipocont  where tcdabr = '" & valor & "'"
    OpenRecordset StrSql, rs_tipocont
    
    If rs_tipocont.EOF Then
        StrSql = " INSERT INTO tipocont(tcdabr,estrnro,tcind) VALUES('" & valor & "'," & CodEstr & ",-1)"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = " SELECT MAX(tcnro) AS CodTC FROM tipocont "
        'StrSql = " SELECT @@IDENTITY AS CodTC "
        OpenRecordset StrSql, rs_TC
        
        StrSql = " INSERT INTO replica_estr(origen,estrnro) VALUES (" & rs_TC!CodTC & "," & CodEstr & ")"
        objConn.Execute StrSql, , adExecuteNoRecords

    End If
End Sub

Private Sub Complementos19(CodEstr As Integer)

    StrSql = " INSERT INTO convenios(estrnro) VALUES(" & CodEstr & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Private Sub Complementos22(CodTer As Integer, CodEstr As Integer, valor As String)

    StrSql = " INSERT INTO formaliq(estrnro,folisistema) VALUES(" & CodEstr & ",-1)"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Private Sub Complementos23(CodTer As Integer, CodEstr As Integer, valor As String)

Dim rs_pos As New ADODB.Recordset

    ' Hay que ver la relacion entra la Osocial y el Plan

    StrSql = " INSERT INTO planos(plnom,osocial) VALUES('" & valor & "'," & CodTer & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    StrSql = " SELECT MAX(plnro) AS CodPl FROM planos "
    'StrSql = " SELECT @@IDENTITY AS CodPl "
    OpenRecordset StrSql, rs_pos
    
    StrSql = " INSERT INTO replica_estr(origen,estrnro) VALUES (" & rs_pos!CodPl & "," & CodEstr & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    

End Sub

Private Sub Complementos40(CodEstr As Integer, CodTer As Integer, valor As String)


    StrSql = " INSERT INTO seguro(ternro,estrnro,segdesc,segest) VALUES(" & CodEstr & "," & CodTer & ",'" & valor & "',-1)"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Private Sub Complementos41(CodEstr As Integer, CodTer As Integer, valor As String)

    StrSql = " INSERT INTO banco(ternro,estrnro,bansucdesc,banest) VALUES(" & CodEstr & "," & CodTer & ",'" & valor & "',-1)"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Private Sub CreaTercero(TipoTer As Integer, valor As String, ByRef CodTer)

Dim rs As New ADODB.Recordset
Dim d_estructura As String
Dim l_pos1 As Integer
Dim l_pos2 As Integer

    
  d_estructura = valor
    
  StrSql = " INSERT INTO tercero(terrazsoc,tersex)"
  StrSql = StrSql & " VALUES('" & d_estructura & "',-1)"
  objConn.Execute StrSql, , adExecuteNoRecords

  StrSql = "SELECT MAX(ternro) AS MaxTernro FROM tercero"
  'StrSql = " SELECT @@IDENTITY AS MaxTernro "
  OpenRecordset StrSql, rs

  CodTer = rs!MaxTernro

  StrSql = " INSERT INTO ter_tip(ternro,tipnro) "
  StrSql = StrSql & " VALUES(" & CodTer & "," & TipoTer & ")"
  objConn.Execute StrSql, , adExecuteNoRecords


End Sub

Private Sub ValidaEstructuraCodExt(TipoEstr As Integer, ByRef valor As String, ByRef CodEst As Integer, ByRef Inserto_estr As Boolean)

Dim Rs_Estr As New ADODB.Recordset

Dim d_estructura As String
Dim CodExt As String
Dim l_pos1 As Byte
Dim l_pos2 As Byte


    d_estructura = valor
    StrSql = " SELECT estrnro FROM estructura WHERE estructura.estrcodext = '" & Left(valor, 20) & "'"
    StrSql = StrSql & " AND estructura.tenro = " & TipoEstr
    OpenRecordset StrSql, Rs_Estr
        
    If Not Rs_Estr.EOF Then
            CodEst = Rs_Estr!estrnro
            Inserto_estr = False
    Else
            StrSql = " INSERT INTO estructura(tenro,estrdabr,empnro,estrest,estrcodext)"
            StrSql = StrSql & " VALUES(" & TipoEstr & ",'" & d_estructura & "',1,-1,'" & Left(d_estructura, 20) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            StrSql = " SELECT MAX(estrnro) AS MaxEst FROM Estructura "      ' Oracle
            'StrSql = " SELECT @@IDENTITY AS MaxEst "                    ' SQL
            OpenRecordset StrSql, Rs_Estr
            
            CodEst = Rs_Estr!MaxEst
            Inserto_estr = True
    End If

End Sub

Private Sub ValidaCategoria(TipoEstr As Integer, ByRef valor As String, nroConv As Integer, ByRef CodEst As Integer, ByRef Inserto_estr As Boolean)
Dim pos1 As Integer
Dim pos2 As Integer

Dim Rs_Estr As New ADODB.Recordset
Dim Rs_Conv As New ADODB.Recordset

Dim d_estructura As String
Dim l_pos1 As Byte
Dim l_pos2 As Byte
Dim CodExt As String

    If InStr(1, valor, "$") > 0 Then
        l_pos1 = InStr(1, valor, "$")
        l_pos2 = Len(valor)
    
        d_estructura = Mid(valor, l_pos1 + 2, l_pos2)
        If l_pos1 <> 0 Then
            CodExt = Mid(valor, 1, l_pos1 - 1)
        Else
            CodExt = ""
        End If
    Else
        d_estructura = valor
        CodExt = ""
    End If
    
    valor = d_estructura
    
'    l_pos1 = InStr(pos1, Valor, "$")
'    l_pos2 = Len(Valor)
'
'    d_estructura = Mid(Valor, pos1 + 1, pos2)
'    CodExt = Mid(Valor, 1, l_pos1 - 1)
'
'    Valor = d_estructura
    
    StrSql = " SELECT estrnro FROM estructura WHERE estructura.estrdabr = '" & d_estructura & "'"
    StrSql = StrSql & " AND estructura.tenro = " & TipoEstr
    OpenRecordset StrSql, Rs_Estr
        
    If Not Rs_Estr.EOF Then
                
          StrSql = " SELECT convnro FROM categoria WHERE categoria.estrnro = " & Rs_Estr!estrnro
          OpenRecordset StrSql, Rs_Conv
                
          If (Not Rs_Conv.EOF) And (nroConv = Rs_Conv!convnro) Then
            
            CodEst = Rs_Estr!estrnro
            Inserto_estr = False
                
          Else
            StrSql = " INSERT INTO estructura(tenro,estrdabr,empnro,estrest,estrcodext)"
            StrSql = StrSql & " VALUES(" & TipoEstr & ",'" & d_estructura & "',1,-1," & CodExt & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(estrnro) AS MaxEst FROM estructura "      'Oracle
            'StrSql = " SELECT @@IDENTITY AS MaxEst "                       'SQL
            OpenRecordset StrSql, Rs_Estr
            
            CodEst = Rs_Estr!MaxEst
            Inserto_estr = True
            
            StrSql = " INSERT INTO categoria(estrnro,convnro) VALUES(" & CodEst & "," & nroConv & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
          End If
                
            
    Else
            StrSql = " INSERT INTO estructura(tenro,estrdabr,empnro,estrest,estrcodext)"
            StrSql = StrSql & " VALUES(" & TipoEstr & ",'" & d_estructura & "',1,-1," & CodExt & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(estrnro) AS MaxEst FROM estructura "      'Oracle
            'StrSql = " SELECT @@IDENTITY AS MaxEst "                        'SQL
            OpenRecordset StrSql, Rs_Estr
            
            CodEst = Rs_Estr!MaxEst
            Inserto_estr = True
            
            StrSql = " INSERT INTO categoria(estrnro,convnro) VALUES(" & CodEst & "," & nroConv & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
    End If


End Sub

Private Sub AsignarEstructura(TipoEstr As Integer, CodEst As Integer, CodTer As Integer, FAlta As String, FBaja As String)

    If CodEst <> 0 Then
    
        If Not FBaja = "Null" Then
    
            StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde,htethasta) VALUES("
            StrSql = StrSql & CodTer & "," & CodEst & "," & TipoEstr & "," & FAlta & "," & FBaja & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde) VALUES("
            StrSql = StrSql & CodTer & "," & CodEst & "," & TipoEstr & "," & FAlta & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
    End If

End Sub

Public Sub LineaModelo_112(ByVal strReg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de Estructuras
' Autor      : FGZ
' Fecha      : 21/09/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Legajo          As String   'LEGAJO                        -- empleado.empleg
Dim Estructura      As String   'Estructura                    -- his_estructura.estrnro
Dim NroTenro        As String   'Tipo de Estructura            -- his_estructura.tenro
Dim FAlta           As String   'Fecha Desde en la Estructura  -- his_estructura.htetdesde
Dim FBaja           As String   'Fecha Hasta en la Estructura  -- his_estructura.htethasta

Dim ternro As Long

Dim pos1 As Integer
Dim pos2 As Integer

Dim NroTercero          As Integer
Dim NroLegajo           As Integer
Dim nro_estructura      As Integer
Dim F_Alta              As String
Dim F_Baja              As String

Dim Inserto_estr        As Boolean

Dim rs As New ADODB.Recordset
Dim rs_sql As New ADODB.Recordset

' True indica que se hace por Descripcion. False por Codigo Externo

Dim EstrDesc             As Boolean   'Sucursal                 -- his_estructura


    EstrDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo

    ' Recupero los Valores del Archivo
    
    pos1 = 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    NroTenro = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Legajo = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Estructura = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    FAlta = Mid(strReg, pos1, pos2 - pos1)
    
    If FAlta = "No se utiliza" Or FAlta = "" Then
        F_Alta = "Null"
    Else
        F_Alta = ConvFecha(FAlta)
    End If
    
    pos1 = pos2 + 1
    pos2 = Len(strReg) + 1
    FBaja = Mid(strReg, pos1, pos2 - pos1)
    
    If FBaja = "No se utiliza" Or FBaja = "" Then
        F_Baja = "Null"
    Else
        F_Baja = ConvFecha(FBaja)
    End If
    
    ' Valida que los campos obligatorios este cargados
    
    If NroTenro = "" Or Legajo = "" Or Estructura = "" Or FAlta = "" Then
        Exit Sub
    End If
    
    ' Busca el Tercero
    StrSql = "SELECT ternro FROM empleado WHERE empleado.empleg = " & Legajo
    OpenRecordset StrSql, rs
    
    If rs.EOF Then Exit Sub
    
    NroTercero = rs!ternro

    ' Validacion y Creacion de la Sucursal (junto con sus Complementos)
    If Estructura <> "No se utiliza" Then
        If EstrDesc Then
            Call ValidaEstructura(CInt(NroTenro), Estructura, nro_estructura, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(CInt(NroTenro), Estructura, nro_estructura, Inserto_estr)
        End If
    End If
    
  ' Inserto las Estructuras
  Call AsignarEstructura(CInt(NroTenro), nro_estructura, NroTercero, F_Alta, F_Baja)
         
  If rs.State = adStateOpen Then
      rs.Close
  End If
End Sub

Private Function TraerCodTipoDocumento(Sigla As String)
    If Not EsNulo(Sigla) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT tidnro FROM tipodocu WHERE tidsigla = '" & Sigla & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO tipodocu (tidsigla) VALUES('"
            StrSql = StrSql & Sigla & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(tidnro) AS Maxtidnro FROM tipodocu "
            OpenRecordset StrSql, rs_sub
                
            TraerCodTipoDocumento = rs_sub!Maxtidnro
        Else
            TraerCodTipoDocumento = rs_sub!tidnro
        End If
    End If
End Function
Private Function TraerCodLocalidad(Localidad As String)
    If Not EsNulo(Localidad) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT locnro FROM localidad WHERE locdesc = '" & Localidad & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO localidad (locdesc) VALUES('"
            StrSql = StrSql & Localidad & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(locnro) AS Maxlocnro FROM localidad "
            OpenRecordset StrSql, rs_sub
                
            TraerCodLocalidad = rs_sub!Maxlocnro
        Else
            TraerCodLocalidad = rs_sub!locnro
        End If
    End If
End Function
Private Function TraerCodProvincia(Provincia As String)
    If Not EsNulo(Provincia) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT provnro FROM Provincia WHERE provdesc = '" & Provincia & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO Provincia (provdesc) VALUES('"
            StrSql = StrSql & Provincia & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(provnro) AS Maxprovnro FROM Provincia "
            OpenRecordset StrSql, rs_sub
                
            TraerCodProvincia = rs_sub!Maxprovnro
        Else
            TraerCodProvincia = rs_sub!provnro
        End If
    End If
End Function
Private Function TraerCodZona(Zona As String)
    If Not EsNulo(Zona) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT zonanro FROM Zona WHERE zonadesc = '" & Zona & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO zona (zonadesc) VALUES('"
            StrSql = StrSql & Zona & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(zonanro) AS Maxzonanro FROM zona "
            OpenRecordset StrSql, rs_sub
                
            TraerCodZona = rs_sub!Maxzonanro
        Else
            TraerCodZona = rs_sub!zonanro
        End If
    End If
End Function
Private Function TraerCodPais(Paisdesc As String)
    If Not EsNulo(Paisdesc) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT paisnro FROM Pais WHERE paisdesc = '" & Paisdesc & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO Pais (paisdesc) VALUES('"
            StrSql = StrSql & Paisdesc & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(paisnro) AS Maxpaisnro FROM pais "
            OpenRecordset StrSql, rs_sub
                
            TraerCodPais = rs_sub!Maxpaisnro
        Else
            TraerCodPais = rs_sub!paisnro
        End If
    End If
End Function
Private Function TraerCodNivelEstudio(Nivdesc As String)
    If Not EsNulo(Nivdesc) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT nivnro FROM nivest WHERE nivdesc = '" & Nivdesc & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            
            StrSql = " INSERT INTO nivest (nivdesc,nivsist,nivobligatorio,nivestfli) VALUES ("
            StrSql = StrSql & "'" & Nivdesc & "'" & ",0,0,0 )"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(nivnro) AS Maxnivnro FROM nivest "
            OpenRecordset StrSql, rs_sub
                
            TraerCodNivelEstudio = CInt(rs_sub!Maxnivnro)
        Else
            TraerCodNivelEstudio = CInt(rs_sub!Nivnro)
        End If
    End If
End Function

Private Function TraerCodTitulo(Titdesabr As String, Nivnro As Integer)
    If Not EsNulo(Titdesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT titnro FROM titulo WHERE titdesabr = '" & Titdesabr & "'"
        StrSql = StrSql & " AND nivnro = " & Nivnro
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO nivest (titdesabr, nivnro ) "
            StrSql = StrSql & " VALUES('" & Titdesabr & "'," & Nivnro & ")"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(titnro) AS Maxtitnro FROM titulo "
            OpenRecordset StrSql, rs_sub
                
            TraerCodTitulo = CInt(rs_sub!Maxtitnro)
        Else
            TraerCodTitulo = CInt(rs_sub!titnro)
        End If
    End If
End Function
'Private Function TraerCodInstitucion(Instdes As String)
'    If Not EsNulo(Instdes) Then
'        Dim rs_sub As New ADODB.Recordset
'        StrSql = " SELECT instnro FROM institucion WHERE instdes = '" & Instdes & "'"
'        OpenRecordset StrSql, rs_sub
'        If rs_sub.EOF Then
'            Dim arreglo As String
'            Dim Cadena As String
'            Dim a As Integer
'            arreglo = Split(Instdes)
'            If UBound(arreglo) >= 1 Then
'                Cadena = Left(Trim(arreglo(a)), 3)
'            Else
'                For a = 0 To UBound(areglo)
'                    Cadena = Cadena & Left(Trim(arreglo(a)), 1)
'                Next a
'            End If
'            StrSql = " INSERT INTO institucion (instdes,instabre, instedu) "
'            StrSql = StrSql & "  VALUES ('" & Instdes & "'" & Cadena & "',-1)"
'
'            objConn.Execute StrSql, , adExecuteNoRecords
'
'            StrSql = " SELECT MAX(nivnro) AS Maxnivnro FROM nivest "
'            OpenRecordset StrSql, rs_sub
'
'            TraerCodInstitucion = CInt(rs_sub!Maxnivnro)
'        Else
'            TraerCodInstitucion = CInt(rs_sub!Nivnro)
'        End If
'    End If
'End Function

'Private Function TraerCodCargo(Cardesabr As String)
'    If Not EsNulo(Cardesabr) Then
'        Dim rs_sub As New ADODB.Recordset
'        StrSql = " SELECT varnro FROM cargo WHERE cardesabr = '" & Cardesabr & "'"
'        OpenRecordset StrSql, rs_sub
'        If rs_sub.EOF Then
'            StrSql = "INSERT INTO cargo (cardesabr ) "
'            StrSql = StrSql & " VALUES('" & Cardesabr & ")"
'
'            objConn.Execute StrSql, , adExecuteNoRecords
'
'            StrSql = " SELECT MAX(carnro) AS Maxcarnro FROM cargo "
'            OpenRecordset StrSql, rs_sub
'
'            TraerCodTitulo = CInt(rs_sub!Maxcarnro)
'        Else
'            TraerCodTitulo = CInt(rs_sub!carnro)
'        End If
'    Else
'        StrSql = "INSERT INTO cargo (cardesabr ) "
'        StrSql = StrSql & " VALUES('" & Cardesabr & ")"
'
'        objConn.Execute StrSql, , adExecuteNoRecords
'
'        StrSql = " SELECT MAX(carnro) AS Maxcarnro FROM cargo "
'        OpenRecordset StrSql, rs_sub
'
'        TraerCodTitulo = CInt(rs_sub!Maxcarnro)
'    End If
'End Function
'
'Private Function TraerCodEmpresa(empdesc As String)
''    Dim Rs_Estr As New ADODB.Recordset
''
''    Dim d_estructura As String
''    Dim CodExt As String
''    Dim l_pos1 As Byte
''    Dim l_pos2 As Byte
''
''    d_estructura = valor
''    StrSql = " SELECT estrnro FROM estructura WHERE estructura.estrcodext = '" & Left(valor, 20) & "'"
''    StrSql = StrSql & " AND estructura.tenro = " & TipoEstr
''    OpenRecordset StrSql, Rs_Estr
''
''    If Not Rs_Estr.EOF Then
''            CodEst = Rs_Estr!estrnro
''            Inserto_estr = False
''    Else
''            StrSql = " INSERT INTO estructura(tenro,estrdabr,empnro,estrest,estrcodext)"
''            StrSql = StrSql & " VALUES(" & TipoEstr & ",'" & d_estructura & "',1,-1,'" & Left(d_estructura, 20) & "')"
''            objConn.Execute StrSql, , adExecuteNoRecords
''            StrSql = " SELECT MAX(estrnro) AS MaxEst FROM Estructura "      ' Oracle
''            'StrSql = " SELECT @@IDENTITY AS MaxEst "                    ' SQL
''            OpenRecordset StrSql, Rs_Estr
''
''            CodEst = Rs_Estr!MaxEst
''            Inserto_estr = True
''    End If
'
'End Function
