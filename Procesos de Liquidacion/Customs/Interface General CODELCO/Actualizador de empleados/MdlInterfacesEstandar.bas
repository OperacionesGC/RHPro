Attribute VB_Name = "MdlInterfacesEstandar"
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



Public Sub Insertar_Linea_Segun_Modelo_Estandar(ByVal Linea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento llamador de acurdo al modelo
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
MyBeginTrans
    Flog.Writeline Espacios(Tabulador * 1) & "Comienza Transaccion"

    HuboError = False
    
    Select Case NroModelo
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
'    Case 216: 'Acumuladores de Agencia para Citrusvil
'        Call LineaModelo_216(Linea)
    Case 217: 'Vales
        Call LineaModelo_217(Linea)
    Case 226: '
'        Call LineaModelo_226(Linea)
    Case 227: '
        Call LineaModelo_227(Linea)
    Case 228: '
        'Call LineaModelo_228(Linea)
        'Modelo reservado para el reporte de Declaracion Jurada
    Case 229: 'Prestamos
        Call LineaModelo_229(Linea)
    Case 230: 'Interface de Dias Pedidos de Vacaciones
        Call LineaModelo_230(Linea)
    Case 231: 'Interface Cta Banco Nacion
        Call LineaModelo_231(Linea)
    Case 232: 'Interface Postulantes Bumerang
        'Reservado en otro proceso
    Case 233: 'Interface Licencias
        Call LineaModelo_233(Linea)
    Case 234: 'Exportacion JDE
        'Reservado en otro proceso
    Case 235: 'Interface de Estadisticas de Accidentes
        Call LineaModelo_235(Linea)
    Case 236: 'IMPORTACION DE Totales de Cantidad de BULTOS  a  RH Pro
        Call LineaModelo_236(Linea)
    Case 237: 'IMPORTACION DE Detalle de Cantidad de BULTOS  a  RH Pro
        Call LineaModelo_237(Linea)
    Case 243: 'Interface Cuentas Corrientes Estandar
        Call LineaModelo_243(Linea)
    Case 244: 'Borrador Detallado
        'Reservado en otro proceso
    Case 245: 'Novedades de Ajuste
        Call LineaModelo_245(Linea)
    Case 246: 'Escala
        Call LineaModelo_246(Linea)
    Case 247: 'Interfase Acumulado de Horas TELEPERFORMANCE
        'esta en el modulo de customizaciones
    Case 248: 'infotipos Deloitte
        'Reservado en otro proceso
    Case 249: 'Interfase de Mapeos para Infotipos
        'Reservado en otro proceso
    Case 250: 'Interfase de Acumuladores del Mes (acu_mes)
        Call LineaModelo_250(Linea)
    End Select

MyCommitTrans
If Not HuboError Then
    'MyCommitTrans
    Flog.Writeline Espacios(Tabulador * 1) & "Transaccion Cometida"
Else
    'MyRollbackTrans
    Flog.Writeline Espacios(Tabulador * 1) & "Transaccion Abortada"
End If
End Sub


Public Sub LineaModelo_211(ByVal strlinea As String)
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
Dim pos1 As Long
Dim pos2 As Long
    
Dim Tercero As Long
Dim NroLegajo As Long

Dim concnro As Long
Dim Conccod As Long
Dim fornro As Long

Dim tpanro As Long
Dim Monto As Single
Dim FechaDesde As String
Dim FechaHasta As String

Dim PeriodoDesde
Dim PeriodoHasta
Dim TieneVigencia As Boolean
Dim EsRetroactivo As Boolean

Dim aux As String
Dim Encontro As Boolean

Dim rs_Empleado As New ADODB.Recordset
Dim rs_Concepto As New ADODB.Recordset
Dim rs_TipoPar As New ADODB.Recordset
Dim rs_NovEmp As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_con_for_tpa As New ADODB.Recordset

' El formato es:
' Formato 1
' Legajo; conccod; tpanro; monto
' ó
' Formato 2.1
' Legajo; conccod; tpanro; monto; FechaDesde; FechaHasta
' ó
' Formato 2.2
' Legajo; conccod; tpanro; monto; FechaDesde
' ó
' Formato 3
' Legajo; conccod; tpanro; monto; MarcaRetroactividad;PeriodoDesde(pliqnro); PeriodoHasta(pliqnro)
' ó
' Formato 4
' Legajo; conccod; tpanro; monto; FechaDesde; FechaHasta; MarcaRetroactividad;PeriodoDesde(pliqnro); PeriodoHasta(pliqnro)
    
    On Error GoTo Manejador_De_Error
    
    TieneVigencia = False
    EsRetroactivo = False

    'Nro de Legajo
    pos1 = 1
    pos2 = InStr(pos1, strlinea, Separador)
    If IsNumeric(Mid$(strlinea, pos1, pos2 - pos1)) Then
        NroLegajo = Mid$(strlinea, pos1, pos2 - pos1)
    Else
        Flog.Writeline Espacios(Tabulador * 1) & "El legajo no es numerico "
        FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": El legajo no es numerico"
        InsertaError 1, 8
        HuboError = True
        Exit Sub
    End If
    
    'Concepto
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Conccod = Mid(strlinea, pos1, pos2 - pos1)

    'Tipo de Parametro
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    tpanro = Mid(strlinea, pos1, pos2 - pos1)

    'Monto
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    If pos2 = 0 Then
        pos2 = Len(strlinea)
        Monto = Mid(strlinea, pos1, pos2)
        Monto = CSng(Replace(CStr(Monto), SeparadorDecimal, "."))
    Else
        Monto = Mid(strlinea, pos1, pos2 - pos1)
        Monto = CSng(Replace(CStr(Monto), SeparadorDecimal, "."))
        
        'Puede veniar Fecha Desde; fecha Hasta ó Retroactivo, Periodo desde , Periodo Hasta
        pos1 = pos2 + 1
        pos2 = InStr(pos1 + 1, strlinea, Separador)
        If pos2 > 0 Then
            aux = Mid(strlinea, pos1, pos2 - pos1)
            If IsDate(aux) Then
                TieneVigencia = True
                'Fecha desde
                FechaDesde = Mid(strlinea, pos1, pos2 - pos1)
            
                'Fecha Hasta
                pos1 = pos2 + 1
                pos2 = InStr(pos1, strlinea, Separador)
                If pos2 > 0 Then
                    FechaHasta = Mid(strlinea, pos1, pos2 - pos1)
                    If IsDate(FechaHasta) Then
                        FechaHasta = CDate(FechaHasta)
                    Else
                        If Not EsNulo(FechaHasta) Then
                            Flog.Writeline Espacios(Tabulador * 1) & "Fecha no valida "
                            FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": La fecha no es valida "
                            InsertaError 1, 4
                            HuboError = True
                            Exit Sub
                        End If
                    End If
                    'Marca de Retroactividad
                    pos1 = pos2 + 1
                    pos2 = InStr(pos1 + 1, strlinea, Separador)
                    aux = Mid(strlinea, pos1, pos2 - pos1)
                    If UCase(aux) = "SI" Then
                        EsRetroactivo = True
                    Else
                        EsRetroactivo = False
                    End If
                
                    'Periodo desde
                    pos1 = pos2 + 1
                    pos2 = InStr(pos1 + 1, strlinea, Separador)
                    PeriodoDesde = Mid(strlinea, pos1, pos2 - pos1)
                    
                    'Periodo hasta
                    pos1 = pos2 + 1
                    pos2 = Len(strlinea)
                    PeriodoHasta = Mid(strlinea, pos1, pos2)
                Else
                    pos2 = Len(strlinea)
                    FechaHasta = Mid(strlinea, pos1, pos2)
                
                    TieneVigencia = True
                End If
            Else
                If UCase(aux) = "SI" Then
                    EsRetroactivo = True
                Else
                    EsRetroactivo = False
                End If
                
                'Periodo desde
                pos1 = pos2 + 1
                pos2 = InStr(pos1 + 1, strlinea, Separador)
                PeriodoDesde = Mid(strlinea, pos1, pos2 - pos1)
                
                'Periodo hasta
                pos1 = pos2 + 1
                pos2 = Len(strlinea)
                PeriodoHasta = Mid(strlinea, pos1, pos2)
            End If
        Else
            'Viene Vigencia con fecha desde y sin fecha hasta
            pos2 = Len(strlinea)
            FechaDesde = Mid(strlinea, pos1, pos2)
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
    Flog.Writeline Espacios(Tabulador * 1) & "No se encontro el legajo " & NroLegajo
    FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": No se encontro el legajo " & NroLegajo
    InsertaError 1, 8
    HuboError = True
    Exit Sub
Else
    Tercero = rs_Empleado!ternro
End If

'Que exista el concepto
StrSql = "SELECT * FROM concepto WHERE conccod = " & Conccod
StrSql = StrSql & " OR conccod = '" & Conccod & "'"
OpenRecordset StrSql, rs_Concepto
If rs_Concepto.EOF Then
    Flog.Writeline Espacios(Tabulador * 1) & "No se encontro el Concepto " & Conccod
    FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": No se encontro el Concepto " & Conccod
    InsertaError 2, 10
    HuboError = True
    Exit Sub
Else
    concnro = rs_Concepto!concnro
    fornro = rs_Concepto!fornro
End If

'Que exista el tipo de Parametro
StrSql = "SELECT * FROM tipopar WHERE tpanro = " & tpanro
OpenRecordset StrSql, rs_TipoPar

If rs_TipoPar.EOF Then
    Flog.Writeline Espacios(Tabulador * 1) & "No se encontro el Tipo de Parametro " & tpanro
    FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": No se encontro el Tipo de Parametro " & tpanro
    InsertaError 3, 11
    HuboError = True
    Exit Sub
End If


'FGZ - 27/01/2004
'Faltaria revisar que el par concepto-parametro se resuelva por novedad
StrSql = "SELECT * FROM con_for_tpa "
StrSql = StrSql & " WHERE concnro = " & concnro
StrSql = StrSql & " AND fornro =" & fornro
StrSql = StrSql & " AND tpanro =" & tpanro
OpenRecordset StrSql, rs_con_for_tpa

If rs_con_for_tpa.EOF Then
    Flog.Writeline Espacios(Tabulador * 1) & "El parametro " & tpanro & " no esta asociado a la formula del concepto " & Conccod
    FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": El parametro " & tpanro & " no esta asociado a la formula del concepto " & Conccod
    InsertaError 3, 11
    HuboError = True
    Exit Sub
Else
    Encontro = False
    Do While Not Encontro And Not rs_con_for_tpa.EOF
        If Not CBool(rs_con_for_tpa!cftauto) Then
            Encontro = True
        End If
        rs_con_for_tpa.MoveNext
    Loop
    If Not Encontro Then
        Flog.Writeline Espacios(Tabulador * 1) & "El parametro " & tpanro & " del concepto " & Conccod & " no se resuelve por novedad "
        FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": El parametro " & tpanro & " del concepto " & Conccod & " no se resuelve por novedad "
        InsertaError 3, 11
        HuboError = True
        Exit Sub
    End If
End If

If EsRetroactivo Then
    'Chequeo que los periodos sean validos
    'Chequeo Periodo Desde
    StrSql = "SELECT * FROM periodo WHERE pliqnro = " & PeriodoDesde
    If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
    OpenRecordset StrSql, rs_Periodo
    
    If rs_Periodo.EOF Then
        Flog.Writeline Espacios(Tabulador * 1) & "Periodo Desde Invalido " & PeriodoDesde
        FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": Periodo Desde Invalido " & PeriodoDesde
        InsertaError 6, 36
        HuboError = True
        Exit Sub
    End If
    
    'Chequeo Periodo Hasta
    StrSql = "SELECT * FROM periodo WHERE pliqnro = " & PeriodoHasta
    If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
    OpenRecordset StrSql, rs_Periodo
    
    If rs_Periodo.EOF Then
        Flog.Writeline Espacios(Tabulador * 1) & "Periodo Hasta Invalido " & PeriodoHasta
        FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": Periodo Hasta Invalido " & PeriodoHasta
        InsertaError 7, 36
        HuboError = True
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
            Flog.Writeline Espacios(Tabulador * 1) & "No se puede insertar la novedad poqrue ya existe una con Vigencia"
            FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": No se puede insertar la novedad poqrue ya existe una con Vigencia"
            InsertaError 1, 94
            HuboError = True
            Exit Sub
        Else
            'Existe una novedad pero sin vigencia ==> Actualizo
            If PisaNovedad Then 'Actualizo la Novedad
                If Not EsRetroactivo Then
                    StrSql = "UPDATE novemp SET nevalor = " & Monto & _
                             " WHERE concnro = " & concnro & _
                             " AND tpanro = " & tpanro & _
                             " AND empleado = " & Tercero
                Else
                    StrSql = "UPDATE novemp SET nevalor = " & Monto & _
                             " , nepliqdesde =  " & PeriodoDesde & _
                             " , nepliqhasta =  " & PeriodoHasta & _
                             " WHERE concnro = " & concnro & _
                             " AND tpanro = " & tpanro & _
                             " AND empleado = " & Tercero
                End If
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.Writeline Espacios(Tabulador * 1) & "Novedad Actualizada "
            Else
                Flog.Writeline Espacios(Tabulador * 1) & "No se insertó la novedad. Ya existe y no se pisa."
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
                     "empleado,concnro,tpanro,nevalor,nevigencia,nepliqdesde,nepliqhasta " & _
                     ") VALUES (" & Tercero & _
                     "," & concnro & _
                     "," & tpanro & _
                     "," & Monto & _
                     "," & CInt(TieneVigencia) & _
                     "," & PeriodoDesde & _
                     "," & PeriodoHasta & _
                     " )"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.Writeline Espacios(Tabulador * 1) & "Novedad insertada "
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
        Flog.Writeline Espacios(Tabulador * 1) & "No se puede insertar la novedad, las vigencias se superponen"
        FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": No se puede insertar la novedad, las vigencias se superponen"
        InsertaError 1, 95
        HuboError = True
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
            StrSql = StrSql & ",nepliqdesde,nepliqhasta"
            StrSql = StrSql & ") VALUES (" & Tercero
            StrSql = StrSql & "," & concnro
            StrSql = StrSql & "," & tpanro
            StrSql = StrSql & "," & Monto
            StrSql = StrSql & ",-1"
            StrSql = StrSql & "," & ConvFecha(FechaDesde)
            If Not EsNulo(FechaHasta) Then
                StrSql = StrSql & "," & ConvFecha(FechaHasta)
            End If
            StrSql = StrSql & "," & PeriodoDesde
            StrSql = StrSql & "," & PeriodoHasta
            StrSql = StrSql & " )"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.Writeline Espacios(Tabulador * 1) & "Novedad insertada "
    End If
End If

Fin:
'Cierro todo y libero
If rs_NovEmp.State = adStateOpen Then rs_NovEmp.Close
If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
If rs_TipoPar.State = adStateOpen Then rs_TipoPar.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_con_for_tpa.State = adStateOpen Then rs_con_for_tpa.Close

Set rs_NovEmp = Nothing
Set rs_Empleado = Nothing
Set rs_Concepto = Nothing
Set rs_TipoPar = Nothing
Set rs_Periodo = Nothing
Set rs_con_for_tpa = Nothing
Exit Sub

Manejador_De_Error:
    HuboError = True

    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strlinea
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin
End Sub


Public Sub LineaModelo_214(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta en emp_ticket y posiblemente en emptikdist (si es que hay distribucion).
' Autor      : FGZ
'              El formato es:
'                   Legajo; sigla; Monto; [catidad1; Valor1 ...[catidad5; Valor5]]
' Fecha      : '29/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1 As Long
Dim pos2 As Long
Dim i As Long

Dim Tercero As Long
Dim NroLegajo As Long

Dim Sigla As String
Dim Monto As Single

Dim cant(5) As Long
Dim Valor(5) As Single
Dim TikValnro(5) As Long
Dim MontoCorrecto As Boolean
Dim TickNro As Long
Dim EtikNro As Long

Dim Cantidades As Long '0-5 Dice la cantidad de pares (cant,valor)

Dim rs_Empleado As New ADODB.Recordset
Dim rs_Ticket As New ADODB.Recordset
Dim rs_EMP_Ticket As New ADODB.Recordset
Dim rs_EMP_TikDist As New ADODB.Recordset
Dim rs_Ticket_Valor As New ADODB.Recordset
Dim rs_TikValor As New ADODB.Recordset

On Error GoTo Manejador_De_Error

Cantidades = 0
' El formato es:
' Legajo; sigla; Monto; [catidad1; Valor1 ...[catidad5; Valor5]]

    'Nro de Legajo
    pos1 = 1
    pos2 = InStr(pos1, strlinea, Separador)
    If IsNumeric(Mid$(strlinea, pos1, pos2 - pos1)) Then
        NroLegajo = Mid$(strlinea, pos1, pos2 - pos1)
    Else
        InsertaError 1, 8
        HuboError = True
        Exit Sub
    End If
    
    'Sigla
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Sigla = Mid(strlinea, pos1, pos2 - pos1)

    'Monto
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    If pos2 = 0 Then
        pos2 = Len(strlinea)
        Monto = Mid(strlinea, pos1, pos2)
    Else
        Monto = Mid(strlinea, pos1, pos2 - pos1)
               
        Cantidades = Cantidades + 1
        'Cantidad 1
        pos1 = pos2 + 1
        pos2 = InStr(pos1 + 1, strlinea, Separador)
        cant(Cantidades) = Mid(strlinea, pos1, pos2 - pos1)
        
        'Valor1
        pos1 = pos2 + 1
        pos2 = InStr(pos1 + 1, strlinea, Separador)
        If pos2 = 0 Then
           pos2 = Len(strlinea)
           Valor(Cantidades) = Mid(strlinea, pos1, pos2)
        Else
           Valor(Cantidades) = Mid(strlinea, pos1, pos2 - pos1)
                
            Cantidades = Cantidades + 1
           'Cantidad 2
           pos1 = pos2 + 1
           pos2 = InStr(pos1 + 1, strlinea, Separador)
           cant(Cantidades) = Mid(strlinea, pos1, pos2 - pos1)
                
           'Valor2
            pos1 = pos2 + 1
            pos2 = InStr(pos1 + 1, strlinea, Separador)
            If pos2 = 0 Then
                pos2 = Len(strlinea)
                Valor(Cantidades) = Mid(strlinea, pos1, pos2)
            Else
                Valor(Cantidades) = Mid(strlinea, pos1, pos2 - pos1)
                        
                Cantidades = Cantidades + 1
                'Cantidad 3
                pos1 = pos2 + 1
                pos2 = InStr(pos1 + 1, strlinea, Separador)
                cant(Cantidades) = Mid(strlinea, pos1, pos2 - pos1)
                     
                'Valor3
                 pos1 = pos2 + 1
                 pos2 = InStr(pos1 + 1, strlinea, Separador)
                 If pos2 = 0 Then
                     pos2 = Len(strlinea)
                     Valor(Cantidades) = Mid(strlinea, pos1, pos2)
                 Else
                     Valor(Cantidades) = Mid(strlinea, pos1, pos2 - pos1)
                     
                     Cantidades = Cantidades + 1
                    'Cantidad 4
                    pos1 = pos2 + 1
                    pos2 = InStr(pos1 + 1, strlinea, Separador)
                    cant(Cantidades) = Mid(strlinea, pos1, pos2 - pos1)
                         
                    'Valor4
                     pos1 = pos2 + 1
                     pos2 = InStr(pos1 + 1, strlinea, Separador)
                     If pos2 = 0 Then
                         pos2 = Len(strlinea)
                         Valor(Cantidades) = Mid(strlinea, pos1, pos2)
                     Else
                         Valor(Cantidades) = Mid(strlinea, pos1, pos2 - pos1)
                         
                         Cantidades = Cantidades + 1
                        'Cantidad 5
                        pos1 = pos2 + 1
                        pos2 = InStr(pos1 + 1, strlinea, Separador)
                        cant(Cantidades) = Mid(strlinea, pos1, pos2 - pos1)
                             
                        'Valor5
                         pos1 = pos2 + 1
                         pos2 = Len(strlinea)
                         Valor(Cantidades) = Mid(strlinea, pos1, pos2)
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
    HuboError = True
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
    HuboError = True
    Exit Sub
Else
    TickNro = rs_Ticket!Tiknro
End If

'Que el monto
MontoCorrecto = True
If Not IsNumeric(Monto) Then
    Flog.Writeline "El monto no es numerico " & Monto
    InsertaError 3, 5
    HuboError = True
    Exit Sub
Else
    Select Case Cantidades
    Case 0:
        MontoCorrecto = True
    Case 1:
        If Monto <> (cant(1) * Valor(1)) Then
            MontoCorrecto = False
        End If
    Case 2:
        If Monto <> (cant(1) * Valor(1) + cant(2) * Valor(2)) Then
            MontoCorrecto = False
        End If
    Case 3:
        If Monto <> (cant(1) * Valor(1) + cant(2) * Valor(2) + cant(3) * Valor(3)) Then
            MontoCorrecto = False
        End If
    Case 4:
        If Monto <> (cant(1) * Valor(1) + cant(2) * Valor(2) + cant(3) * Valor(3) + cant(4) * Valor(4)) Then
            MontoCorrecto = False
        End If
    Case 5:
        If Monto <> (cant(1) * Valor(1) + cant(2) * Valor(2) + cant(3) * Valor(3) + cant(4) * Valor(4) + cant(5) * Valor(5)) Then
            MontoCorrecto = False
        End If
    End Select
End If
If Not MontoCorrecto Then
    Flog.Writeline "La suma de los detalles no coincide con el monto " & Monto
    InsertaError 3, 92
    HuboError = True
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
            StrSql = StrSql & " tvalmonto = " & Valor(i)
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
                    HuboError = True
                    Exit Sub
                End If
            Else
                Flog.Writeline "Valor no encontrado en TIKVALOR " & Valor(i)
                InsertaError 3 + Cantidades, 64
                HuboError = True
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
                StrSql = StrSql & "," & Valor(i) * cant(i)
                StrSql = StrSql & "," & Valor(i)
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
            StrSql = StrSql & " tvalmonto = " & Valor(i)
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
                    HuboError = True
                    Exit Sub
                End If
            Else
                Flog.Writeline "Valor no encontrado en TIKVALOR " & Valor(i)
                InsertaError 3 + Cantidades, 64
                HuboError = True
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
            StrSql = StrSql & "," & Valor(i) * cant(i)
            StrSql = StrSql & "," & Valor(i)
            StrSql = StrSql & "," & cant(i)
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
        Next i
        
        Flog.Writeline "Ticket Actualizado "
    Else
        Flog.Writeline "Ticket no actualizado porque es manual o fué liquidado "
    End If
End If

Fin:
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

Exit Sub

Manejador_De_Error:
    HuboError = True

    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strlinea
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin

End Sub


Public Sub LineaModelo_215(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta en Acu_age.
' Autor      : FGZ
'              El formato es:
'                   Legajo; Acunro; Monto; catidad; año,mes
' Fecha      : 29/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1 As Long
Dim pos2 As Long
Dim i As Long

Dim Tercero As Long
Dim NroLegajo As Long
Dim acunro As Long
Dim Monto As Single
Dim Cantidad As Single
Dim Anio As Long
Dim mes As Long
Dim PliqNro As Long

Dim rs_Empleado As New ADODB.Recordset
Dim rs_Acumulador As New ADODB.Recordset
Dim rs_Acu_Age As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset

On Error GoTo Manejador_De_Error

' El formato es:
' Legajo; Acunro; Monto; catidad; año,mes

    'Nro de Legajo
    pos1 = 1
    pos2 = InStr(pos1, strlinea, Separador)
    If IsNumeric(Mid$(strlinea, pos1, pos2 - pos1)) Then
        NroLegajo = Mid$(strlinea, pos1, pos2 - pos1)
    Else
        InsertaError 1, 8
        HuboError = True
        Exit Sub
    End If
    
    'Acumulador
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    acunro = Mid(strlinea, pos1, pos2 - pos1)

    'Monto
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Monto = Mid(strlinea, pos1, pos2 - pos1)
               
    'Cantidad
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Cantidad = Mid(strlinea, pos1, pos2 - pos1)
        
    'Año
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Anio = Mid(strlinea, pos1, pos2 - pos1)
        
    'Mes
    pos1 = pos2 + 1
    pos2 = Len(strlinea)
    mes = Mid(strlinea, pos1, pos2)

' ====================================================================
'   Validar los parametros Levantados

'Que exista el legajo
StrSql = "SELECT * FROM empleado where empleg = " & NroLegajo
OpenRecordset StrSql, rs_Empleado
If rs_Empleado.EOF Then
    Flog.Writeline "No se encontro el legajo " & NroLegajo
    InsertaError 1, 8
    HuboError = True
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
    HuboError = True
    Exit Sub
End If

'Que el monto sea valido
If Not IsNumeric(Monto) Then
    Flog.Writeline "El monto no es numerico " & Monto
    InsertaError 3, 5
    HuboError = True
    Exit Sub
End If

'Que la cantidad sea valida
If Not IsNumeric(Cantidad) Then
    Flog.Writeline "La cantidad no es numerico " & Cantidad
    InsertaError 4, 5
    HuboError = True
    Exit Sub
End If

'Que el año sea valido
If Not IsNumeric(Anio) Then
    Flog.Writeline "El año no es numerico " & Anio
    InsertaError 5, 5
    HuboError = True
    Exit Sub
End If

'Que el mes sea valido
If Not IsNumeric(mes) Then
    Flog.Writeline "El mes no es numerico " & Anio
    InsertaError 6, 5
    HuboError = True
    Exit Sub
End If

'Busco el pliqnro correspondiente a ese año y mes
StrSql = "SELECT * FROM PERIODO WHERE pliqmes =" & mes
StrSql = StrSql & " AND pliqanio =" & Anio
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.Writeline "No existe periodo correspondiente al año " & Anio & " y mes  " & mes
    InsertaError 6, 5
    HuboError = True
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

Fin:
'cierro y libero
If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Acu_Age.State = adStateOpen Then rs_Acu_Age.Close
If rs_Acumulador.State = adStateOpen Then rs_Acumulador.Close

Set rs_Empleado = Nothing
Set rs_Periodo = Nothing
Set rs_Acu_Age = Nothing
Set rs_Acumulador = Nothing
Exit Sub

Manejador_De_Error:
    HuboError = True

    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strlinea
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin

End Sub

Public Sub LineaModelo_217(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta en vales.
'              El formato es:
'                   Legajo; Monto; Fecha ; Tipo
' Autor      : FGZ
' Fecha      : 29/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1 As Long
Dim pos2 As Long
Dim i As Long

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

On Error GoTo Manejador_De_Error
' El formato es:
' Legajo; Monto; Fecha ; Tipo

    'Nro de Legajo
    pos1 = 1
    pos2 = InStr(pos1, strlinea, Separador)
    If IsNumeric(Mid$(strlinea, pos1, pos2 - pos1)) Then
        NroLegajo = Mid$(strlinea, pos1, pos2 - pos1)
    Else
        InsertaError 1, 8
        HuboError = True
        Exit Sub
    End If
    
    'Monto
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Monto = Mid(strlinea, pos1, pos2 - pos1)
               
    'Fecha
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    FechaVale = CDate(Mid(strlinea, pos1, pos2 - pos1))
               
    pos1 = pos2 + 1
    pos2 = Len(strlinea)
    TipoVale = Mid(strlinea, pos1, pos2)

' ====================================================================
'   Validar los parametros Levantados

'Que exista el legajo
StrSql = "SELECT * FROM empleado where empleg = " & NroLegajo
OpenRecordset StrSql, rs_Empleado
If rs_Empleado.EOF Then
    Flog.Writeline "No se encontro el legajo " & NroLegajo
    InsertaError 1, 8
    HuboError = True
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
    HuboError = True
    Exit Sub
Else
    PliqNro = rs_Periodo!PliqNro
End If

'Que el monto
MontoCorrecto = True
If Not IsNumeric(Monto) Then
    Flog.Writeline "El monto no es numerico " & Monto
    InsertaError 2, 5
    HuboError = True
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
    HuboError = True
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

Fin:
'cierro y libero
If rs_TipoVale.State = adStateOpen Then rs_TipoVale.Close
If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close

Set rs_TipoVale = Nothing
Set rs_Empleado = Nothing
Set rs_Periodo = Nothing

Exit Sub

Manejador_De_Error:
    HuboError = True

    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strlinea
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin

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
'Dim pos1            as long
'Dim pos2            as long
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
'    Dim tidnro as long       '(01)Tipo Doc
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
'    Dim locnro as long       '(15)Localidad
'    Dim provnro as long      '(16)Provincia
'    Dim partnro as long      '(17)Partido
'    Dim zonanro as long      '(18)Zona
'    Dim paisnro as long      '(19)Pais
''POSTULANTE
'    Dim terfecnac As Date       '(20)Naciemiento
'    Dim tersex As Boolean       '(21)Sexo
''ESTUDIOS FORMALES 1
'    Dim Nivnro as long       '(22)Nivel
'    Dim capcomp As Boolean      '(23)Completo
'    Dim titnro as long       '(24)Titulo
'    Dim instnro as long      '(25)Institucion
''ESTUDIOS FORMALES1 2
'    Dim nivnro2 as long      '(26)Posgrado
'    Dim titnro2 as long      '(27)Postitulo
'    Dim instnro2 as long     '(28)Institucion
''EXPERIENCIA LABORAL 1
'    Dim carnro as long       '(29)Cargo anterior
'    Dim lempnro as long      '(30)Empresa
'    Dim empatareas As String    '(31)Tarea Desempeñada
'    Dim empaini As Date         '(32)Fec Desde
'    Dim empafin As Date         '(33)Fec Hasta
''EXPERIENCIA LABORAL 2
'    Dim carnro2 as long      '(34)Cargo anterior 2
'    Dim lempnro2 as long     '(35)Empresa
'    Dim empatareas2 As String   '(36)Tarea Desempeñada
'    Dim empaini2 As Date        '(37)Fec Desde
'    Dim empafin2 As Date        '(38)Fec Hasta
''IDIOMAS
'    Dim idinro1 as long      '(39)Idioma 1
'    Dim emidlee1 as long     '(40)Lee: Nivel
'    Dim empidhabla1 as long  '(41)Habla: Nivel
'    Dim empidescr1 as long   '(42)Escribe: Nivel
'    Dim idinro2 as long      '(43)Idioma 2
'    Dim emidlee2 as long     '(44)Lee: Nivel
'    Dim empidhabla2 as long  '(45)Habla: Nivel
'    Dim empidescr2 as long   '(46)Escribe: Nivel
'    Dim idinro3 as long      '(47)Idioma 3
'    Dim emidlee3 as long     '(48)Lee: Nivel
'    Dim empidhabla3 as long  '(49)Habla: Nivel
'    Dim empidescr3 as long   '(50)Escribe: Nivel
''OTROS CURSOS 1
'    Dim estinfdesabr As String  '(51)desc.Curso
'    Dim tipcurnro as long    '(52)Tipo Curso
'    Dim estinffecha As Date     '(53)Fec Curso
'    Dim instnro3 as long     '(54)Institucion
''OTROS CURSOS 2
'    Dim estinfdesabr2 As String '(55)desc.Curso
'    Dim tipcurnro2 as long   '(56)Tipo Curso
'    Dim estinffecha2 As Date    '(57)Fec Curso
'    'Dim instnro3 as long     '(58)Institucion
'
'    Dim ternro as long
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
'Dim NroTercero          as long
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


Public Sub LineaModelo_227(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1            As Long
Dim pos2            As Long

On Error GoTo Manejador_De_Error




Fin:
Exit Sub

Manejador_De_Error:
    HuboError = True

    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strlinea
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin
End Sub

Public Sub LineaModelo_228(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    'FGZ -  22/09/2004
    'No USAR. este modelo se usó en el reporte de Declaracion Jurada de la estrella.
    
On Error GoTo Manejador_De_Error
    
Fin:
Exit Sub

Manejador_De_Error:
    HuboError = True

    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strlinea
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin
    
End Sub

Public Sub LineaModelo_229(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Interface de Prestamos
' Autor      : FGZ
' Fecha      : 27/09/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1            As Long
Dim pos2            As Long

Dim Nro_Legajo As Long
Dim Lin_Prestamo As Long
Dim Descripcion As String
Dim Sucursal As Long
Dim Nro_Comprobante As Long
Dim Cant_Cuotas As Long
Dim Monto_Total As Single
Dim Anio As Long
Dim mes As Long
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
Dim rs_Estructura  As New ADODB.Recordset


On Error GoTo Manejador_De_Error

' El formato es:
'   Legajo; Linea de Prestamo; Descripcion; Sucursal; Nro Comprobante; Cant. de Cuotas;
    'Monto Total; Año; Mes; Fecha Otorg.

    'Nro de Legajo
    pos1 = 1
    pos2 = InStr(pos1, strlinea, Separador)
    If IsNumeric(Mid$(strlinea, pos1, pos2 - pos1)) Then
        Nro_Legajo = Mid$(strlinea, pos1, pos2 - pos1)
    Else
        Flog.Writeline Espacios(Tabulador * 1) & "El legajo no es numerico"
        FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": El legajo no es numerico"
        InsertaError 1, 8
        HuboError = True
        Exit Sub
    End If
    
    'Linea de Prestamo
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Lin_Prestamo = Mid(strlinea, pos1, pos2 - pos1)

    'Descripcion
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Descripcion = Mid(strlinea, pos1, pos2 - pos1)

    'Sucursal
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Sucursal = Mid(strlinea, pos1, pos2 - pos1)

    'Nro Comprobante
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Nro_Comprobante = Mid(strlinea, pos1, pos2 - pos1)

    'Cantidad de Cuotas
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Cant_Cuotas = Mid(strlinea, pos1, pos2 - pos1)

    'Monto Total
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Monto_Total = Mid(strlinea, pos1, pos2 - pos1)
    Monto_Total = CSng(Replace(CStr(Monto_Total), SeparadorDecimal, "."))
    
    'Año
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Anio = Mid(strlinea, pos1, pos2 - pos1)

    'Mes
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    mes = Mid(strlinea, pos1, pos2 - pos1)

    'Fecha de Otorgamiento
    pos1 = pos2 + 1
    pos2 = Len(strlinea)
    Fecha_Otorg = Mid(strlinea, pos1, pos2)

' ----------------------------------
'Validaciones

'Que exista el legajo
StrSql = "SELECT * FROM empleado where empleg = " & Nro_Legajo
OpenRecordset StrSql, rs_Empleado
If rs_Empleado.EOF Then
    Flog.Writeline Espacios(Tabulador * 1) & "No se encontro el legajo " & Nro_Legajo
    FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": No se encontro el legajo " & Nro_Legajo
    InsertaError 1, 8
    HuboError = True
    Exit Sub
Else
    Tercero = rs_Empleado!ternro
End If

'que exista la sucursal
StrSql = "SELECT * FROM sucursal where estrnro = " & Sucursal
OpenRecordset StrSql, rs_Sucursal
If rs_Sucursal.EOF Then
    'Sucursal = 0
    Flog.Writeline Espacios(Tabulador * 1) & "No se encontro la sucursal " & Sucursal
    Flog.Writeline Espacios(Tabulador * 1) & "Busco la sucursal actual"
    'FlogE.writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": No se encontro la sucursal " & Sucursal
    
    'buscar la Sucursal del Empleado
    StrSql = " SELECT estrnro FROM his_estructura " & _
             " WHERE ternro = " & Tercero & " AND " & _
             " tenro = 1 AND " & _
             " (htetdesde <= " & ConvFecha(Fecha_Otorg) & ") AND " & _
             " ((" & ConvFecha(Fecha_Otorg) & " <= htethasta) or (htethasta is null))"
    If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
    OpenRecordset StrSql, rs_Estructura
    
    If Not rs_Estructura.EOF Then
        Sucursal = rs_Estructura!estrnro
    Else
        Flog.Writeline Espacios(Tabulador * 1) & "El empleado no tiene ninguna sucursal a la fecha"
        FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": El empleado no tiene ninguna sucursal a la fecha"
        InsertaError 4, 56
        HuboError = True
        Exit Sub
    End If
End If

'Que el monto sea numerico
If Not IsNumeric(Monto_Total) Then
    Flog.Writeline Espacios(Tabulador * 1) & "El monto no es numerico " & Monto_Total
    FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": El monto no es numerico " & Monto_Total
    InsertaError 7, 5
    HuboError = True
    Exit Sub
End If

'que el mes sea valido
If IsNumeric(mes) Then
    If mes > 12 Or mes < 1 Then
        Flog.Writeline Espacios(Tabulador * 1) & "El mes es incorrecto " & mes
        FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": El mes es incorrecto " & mes
        InsertaError 9, 42
        HuboError = True
        Exit Sub
    End If
Else
    Flog.Writeline Espacios(Tabulador * 1) & "El mes es incorrecto " & mes
    FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": El mes es incorrecto " & mes
    InsertaError 9, 42
    HuboError = True
    Exit Sub
End If

'que el año sea valido
If Not IsNumeric(Anio) Then
    Flog.Writeline Espacios(Tabulador * 1) & "El año es incorrecto " & Anio
    FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": El año es incorrecto " & Anio
    InsertaError 8, 3
    HuboError = True
    Exit Sub
End If

'que la fecha de otorgamiento sea una fecha
If Not IsDate(Fecha_Otorg) Then
    Flog.Writeline Espacios(Tabulador * 1) & "La fecha es incorrecta " & Fecha_Otorg
    FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": La fecha es incorrecta " & Fecha_Otorg
    InsertaError 10, 4
    HuboError = True
    Exit Sub
End If

'Busco el Estado de prestamos (Nuevo)
StrSql = "SELECT * FROM estadopre ORDER BY estnro "
OpenRecordset StrSql, rs_Estado
If rs_Estado.EOF Then
    Flog.Writeline Espacios(Tabulador * 1) & "No se encontro el estado Pendiente para los prestamos "
    FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": No se encontro el estado Pendiente para los prestamos "
    EstNro = 0
Else
    EstNro = rs_Estado!EstNro
End If

'Busco la primer moneda
StrSql = "SELECT * FROM moneda"
OpenRecordset StrSql, rs_Monedas
If rs_Monedas.EOF Then
    Flog.Writeline Espacios(Tabulador * 1) & "No se encontro ninguna Moneda. Moneda Default 0"
    FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": No se encontro ninguna Moneda. Moneda Default 0"
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
    Flog.Writeline Espacios(Tabulador * 1) & "No se encontro periodo de liquidacion asociado"
    FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": No se encontro periodo de liquidacion asociado"
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
    Flog.Writeline Espacios(Tabulador * 1) & "No se encontro la linea de Prestamo"
    FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": No se encontro la linea de Prestamo"
    InsertaError 2, 96
    HuboError = True
    Exit Sub
Else
    lnprenro = rs_pre_linea!lnprenro
End If


' ---------------------------------------------
'Inserto

'inserto en prestamo
StrSql = "INSERT INTO prestamo ("
StrSql = StrSql & "predesc,preimp,precantcuo,ternro,quincenal,premes,preanio,monnro,estnro"
StrSql = StrSql & ",lnprenro,iduser,prefecotor,sucursal,pliqnro,precompr"
StrSql = StrSql & ",prequin,pretna,preiva,preotrosgas,prediavto "
StrSql = StrSql & ") VALUES ("
StrSql = StrSql & "'" & Left(Descripcion, 60) & "'"
StrSql = StrSql & "," & Monto_Total
StrSql = StrSql & "," & Cant_Cuotas
StrSql = StrSql & "," & Tercero
StrSql = StrSql & ",0"
StrSql = StrSql & "," & mes
StrSql = StrSql & "," & Anio
StrSql = StrSql & "," & Moneda
StrSql = StrSql & "," & EstNro

StrSql = StrSql & "," & lnprenro
StrSql = StrSql & ",0"
StrSql = StrSql & "," & ConvFecha(Fecha_Otorg)
StrSql = StrSql & "," & Sucursal
StrSql = StrSql & "," & PliqNro
StrSql = StrSql & ",'" & Nro_Comprobante & "'"
StrSql = StrSql & ",0,0,0,0,1 "
StrSql = StrSql & " )"
objConn.Execute StrSql, , adExecuteNoRecords

Flog.Writeline Espacios(Tabulador * 1) & "Prestamo Insertado"

Fin:
'Cierro y libero
If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
If rs_TipoPrestamo.State = adStateOpen Then rs_TipoPrestamo.Close
If rs_Sucursal.State = adStateOpen Then rs_Sucursal.Close
If rs_Monedas.State = adStateOpen Then rs_Monedas.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Estado.State = adStateOpen Then rs_Estado.Close
If rs_pre_linea.State = adStateOpen Then rs_pre_linea.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close

Set rs_Empleado = Nothing
Set rs_TipoPrestamo = Nothing
Set rs_Sucursal = Nothing
Set rs_Monedas = Nothing
Set rs_Periodo = Nothing
Set rs_Estado = Nothing
Set rs_pre_linea = Nothing
Set rs_Estructura = Nothing
Exit Sub

Manejador_De_Error:
    HuboError = True

    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strlinea
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin

End Sub


'Public Sub LineaModelo_230_old(ByVal strLinea As String)
'' ---------------------------------------------------------------------------------------------
'' Descripcion: Interface de Dias Pedidos de Vacaciones
'' Autor      : FGZ
'' Fecha      : 04/10/2004
'' Ultima Mod.: Pedió vigencia el 21/12/2004
'' Descripcion: cambio de formato
'' ---------------------------------------------------------------------------------------------
'' El formato es:
''   Legajo; Nombre y Apellido; Antiguedad Años; Antiguedad Meses; Dias Pendientes; Dias Correspondientes;
''   Total Dias; Fecha Desde; Fecha Hasta
'' ---------------------------------------------------------------------------------------------
'Dim pos1            as long
'Dim pos2            as long
'
'Dim Nro_Legajo As Long
'Dim Nombre_Apellido As String
'Dim Ant_Anios as long
'Dim Ant_Meses as long
'Dim Aux_Dias_Pendientes as long
'Dim Aux_Dias_Correspondientes as long
'Dim Aux_TipoVacacion As String
'Dim Total_Dias as long
'Dim Fecha_Desde As Date
'Dim Fecha_Hasta As Date
'Dim Tercero As Long
'
'Dim Aux_Fecha_Desde As Date
'
'Dim rs_tipovacac As New ADODB.Recordset
'Dim rs_vacdiascor As New ADODB.Recordset
'Dim rsDias As New ADODB.Recordset
'Dim rsVac As New ADODB.Recordset
'
'Dim diascoract as long
'Dim diastom as long
'Dim diascorant as long
'Dim diasdebe as long
'Dim diastot as long
'Dim diasyaped as long
'Dim diaspend as long
'
'Dim nroTipvac As Long
'Dim Hasta As Date
'Dim totferiados as long
'Dim tothabiles as long
'Dim totNohabiles as long
'Dim NroVac As Long
'
'Dim rs_Empleado As New ADODB.Recordset
'Dim rs_Periodos_Vac As New ADODB.Recordset
'
'
''Actio el Manejador de Errores Local
'On Error GoTo Manejador_De_Error
'
'    'Nro de Legajo
'    pos1 = 1
'    pos2 = InStr(pos1, strLinea, Separador)
'    If IsNumeric(Mid$(strLinea, pos1, pos2 - pos1)) Then
'        Nro_Legajo = Mid$(strLinea, pos1, pos2 - pos1)
'    Else
'        InsertaError 1, 8
'        HuboError = True
'        Exit Sub
'    End If
'
'    'Nombre y Apellido
'    pos1 = pos2 + 1
'    pos2 = InStr(pos1 + 1, strLinea, Separador)
'    Nombre_Apellido = Mid(strLinea, pos1, pos2 - pos1)
'
'    'Antiguedad Años
'    pos1 = pos2 + 1
'    pos2 = InStr(pos1 + 1, strLinea, Separador)
'    Ant_Anios = Mid(strLinea, pos1, pos2 - pos1)
'
'    'Antiguedad Mese
'    pos1 = pos2 + 1
'    pos2 = InStr(pos1 + 1, strLinea, Separador)
'    Ant_Meses = Mid(strLinea, pos1, pos2 - pos1)
'
'    'Tipo de Vacaciones
'    pos1 = pos2 + 1
'    pos2 = InStr(pos1 + 1, strLinea, Separador)
'    Aux_TipoVacacion = Mid(strLinea, pos1, pos2 - pos1)
'
'    'Dias_Pendientes
'    pos1 = pos2 + 1
'    pos2 = InStr(pos1 + 1, strLinea, Separador)
'    Aux_Dias_Pendientes = Mid(strLinea, pos1, pos2 - pos1)
'
'    'Dias Correspondientes
'    pos1 = pos2 + 1
'    pos2 = InStr(pos1 + 1, strLinea, Separador)
'    Aux_Dias_Correspondientes = Mid(strLinea, pos1, pos2 - pos1)
'
'    'Total de Dias
'    pos1 = pos2 + 1
'    pos2 = InStr(pos1 + 1, strLinea, Separador)
'    Total_Dias = Mid(strLinea, pos1, pos2 - pos1)
'
'    'Fecha Desde
'    pos1 = pos2 + 1
'    pos2 = InStr(pos1 + 1, strLinea, Separador)
'    Fecha_Desde = Mid(strLinea, pos1, pos2 - pos1)
'
'    'Fecha Hasta
'    pos1 = pos2 + 1
'    pos2 = Len(strLinea)
'    Fecha_Hasta = Mid(strLinea, pos1, pos2)
'
'    ' ----------------------------------
'    'Validaciones
'
'    'Que exista el legajo
'    StrSql = "SELECT * FROM empleado where empleg = " & Nro_Legajo
'    OpenRecordset StrSql, rs_Empleado
'    If rs_Empleado.EOF Then
'        Flog.writeline "No se encontro el legajo " & Nro_Legajo
'        InsertaError 1, 8
'        HuboError = True
'        Exit Sub
'    Else
'        Tercero = rs_Empleado!Ternro
'    End If
'
'    Aux_Fecha_Desde = Fecha_Desde
'
'    'Busco todos los Periodos Involucrados entre las Fechas
'    StrSql = "SELECT * FROM vacacion "
'    StrSql = StrSql & " WHERE vacfecdesde <= " & ConvFecha(Fecha_Hasta)
'    StrSql = StrSql & " AND  vacfechasta >= " & ConvFecha(Fecha_Desde)
'    StrSql = StrSql & " ORDER BY vacnro"
'    OpenRecordset StrSql, rs_Periodos_Vac
'
'    Do While Not rs_Periodos_Vac.EOF And Aux_Fecha_Desde < Fecha_Hasta
'        'si le fecha en la que se va a generar los dias pedidos estan fuera del rengo de fechas del periodo
'        ' no se procesan
'        If Aux_Fecha_Desde >= rs_Periodos_Vac!vacfecdesde And Aux_Fecha_Desde <= rs_Periodos_Vac!vacfechasta Then
'            diascoract = 0
'            diastom = 0
'            diascorant = 0
'            diasdebe = 0
'            diastot = 0
'            diasyaped = 0
'            diaspend = 0
'
'            Flog.writeline "Periodo de Vacaciones:" & rs_Periodos_Vac!vacnro & " " & rs_Periodos_Vac!vacdesc
'
'            NroVac = rs_Periodos_Vac!vacnro
'
'            StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Tercero & " AND vacnro = " & NroVac
'            OpenRecordset StrSql, rs_vacdiascor
'            If Not rs_vacdiascor.EOF Then
'
'                StrSql = "SELECT * FROM tipovacac WHERE tipvacnro = " & rs_vacdiascor!tipvacnro
'                OpenRecordset StrSql, rs_tipovacac
'                If Not rs_tipovacac.EOF Then
'                    nroTipvac = rs_tipovacac!tipvacnro
'                End If
'
'                diascoract = rs_vacdiascor!vdiascorcant ' dias corresp al periodo actual
'
'                StrSql = "SELECT * FROM vacacion WHERE vacnro <> " & NroVac & " and vacfechasta < " & ConvFecha(Fecha_Desde)
'                OpenRecordset StrSql, rsVac
'                Do While Not rsVac.EOF
'
'                    diastom = 0
'
'                    StrSql = "SELECT * FROM lic_vacacion " & _
'                             " INNER JOIN emp_lic ON emp_lic.emp_licnro = lic_vacacion.emp_licnro " & _
'                             " WHERE lic_vacacion.vacnro = " & rsVac!vacnro & " AND emp_lic.empleado = " & Tercero
'                    OpenRecordset StrSql, rsDias
'                    Do While Not rsDias.EOF
'                        diastom = diastom + rsDias!elcantdias
'                        rsDias.MoveNext
'                    Loop
'                    diascorant = 0
'                    StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Tercero & " AND vacnro = " & rsVac!vacnro
'                    OpenRecordset StrSql, rs
'                    If Not rs.EOF Then diascorant = rs!vdiascorcant
'                    diasdebe = diasdebe + (diascorant - diastom)
'
'                    rsVac.MoveNext
'                Loop
'
'                diastot = diascoract + diasdebe
'            End If
'
'
'            ' Busco los pedidos de ese periodo
'            StrSql = "SELECT * FROM vacdiasped WHERE ternro = " & Tercero & " AND vacnro = " & NroVac
'            OpenRecordset StrSql, objRs
'            Do While Not objRs.EOF
'                'diasyaped = diasyaped + objRs!vdiapedcant
'                diasyaped = diasyaped + objRs!vdiaspedhabiles
'                objRs.MoveNext
'            Loop
'
'            diaspend = diastot - diasyaped
'            If diaspend > 0 Then
'
'                Call DiasPedidos(NroVac, nroTipvac, Aux_Fecha_Desde, Fecha_Hasta, Hasta, Tercero, diaspend, tothabiles, totNohabiles, totferiados)
'
'                StrSql = "INSERT INTO vacdiasped (vdiapedhasta,vdiaspedferiados,vdiaspedhabiles,vdiaspednohabiles,ternro,vacnro,vdiapedcant,vdiapeddesde,vdiaspedestado) VALUES (" & _
'                          ConvFecha(Hasta) & "," & totferiados & "," & tothabiles & "," & totNohabiles & "," & Tercero & "," & NroVac & "," & (totferiados + tothabiles + totNohabiles) & "," & ConvFecha(Aux_Fecha_Desde) & ",-1)"
'                objConn.Execute StrSql, , adExecuteNoRecords
'
'                Aux_Fecha_Desde = Hasta + 1
'            End If
'        Else
'            Flog.writeline "La fecha en la que se va a generar los dias pedidos estan fuera del rengo de fechas del periodo " & Aux_Fecha_Desde
'        End If
'
'        rs_Periodos_Vac.MoveNext
'    Loop
'
'    If Aux_Fecha_Desde < Fecha_Hasta Then
'        Flog.writeline "Quedaron " & DateDiff("d", Fecha_Hasta, Aux_Fecha_Desde) + 1 & " días sin generar "
'    End If
'
'fin:
'
''Desactivo el manejador de Errores Local
'On Error GoTo 0
'
'Exit Sub
'
'Manejador_De_Error:
'    HuboError = True
'
'    Flog.writeline
'    Flog.writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strLinea
'    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
'    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
'    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
'    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
'    Flog.writeline
'    If InStr(1, Err.Description, "ODBC") > 0 Then
'        'Fue error de Consulta de SQL
'        Flog.writeline
'        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
'        Flog.writeline
'    End If
'    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
'    Flog.writeline
'    GoTo fin
'End Sub

Public Sub LineaModelo_230(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Interface de Dias Pedidos de Vacaciones
' Autor      : FGZ
' Fecha      : 04/10/2004
' Ultima Mod.:
' Descripcion: cambio de formato - 21/12/2004
' ---------------------------------------------------------------------------------------------
' El formato es:
'   Legajo; ; Fecha Desde; Cantidad de dias
' ---------------------------------------------------------------------------------------------
Dim pos1            As Long
Dim pos2            As Long

Dim Nro_Legajo As Long
Dim Nombre_Apellido As String
Dim Ant_Anios As Long
Dim Ant_Meses As Long
Dim Aux_Dias_Pendientes As Long
Dim Aux_Dias_Correspondientes As Long
Dim Aux_TipoVacacion As String
Dim Total_Dias As Long
Dim Fecha_Desde As Date
Dim Fecha_Hasta As Date
Dim Tercero As Long

Dim Aux_Fecha_Desde As Date
Dim Aux_Dias As Long

Dim rs_tipovacac As New ADODB.Recordset
Dim rs_vacdiascor As New ADODB.Recordset
Dim rsDias As New ADODB.Recordset
Dim rsVac As New ADODB.Recordset

Dim diascoract As Long
Dim diastom As Long
Dim diascorant As Long
Dim diasdebe As Long
Dim diastot As Long
Dim diasyaped As Long
Dim diaspend As Long

Dim nroTipvac As Long
Dim Hasta As Date
Dim totferiados As Long
Dim tothabiles As Long
Dim totNohabiles As Long
Dim NroVac As Long

Dim rs_Empleado As New ADODB.Recordset
Dim rs_Periodos_Vac As New ADODB.Recordset


'Actio el Manejador de Errores Local
On Error GoTo Manejador_De_Error

    'Nro de Legajo
    pos1 = 1
    pos2 = InStr(pos1, strlinea, Separador)
    If IsNumeric(Mid$(strlinea, pos1, pos2 - pos1)) Then
        Nro_Legajo = Mid$(strlinea, pos1, pos2 - pos1)
    Else
        InsertaError 1, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fecha Desde
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Fecha_Desde = Mid(strlinea, pos1, pos2 - pos1)
    
    'Total de dias Pedidos
    pos1 = pos2 + 1
    pos2 = Len(strlinea)
    Total_Dias = Mid(strlinea, pos1, pos2 - pos1 + 1)

    ' ----------------------------------
    'Validaciones
    
    'Que exista el legajo
    StrSql = "SELECT * FROM empleado where empleg = " & Nro_Legajo
    OpenRecordset StrSql, rs_Empleado
    If rs_Empleado.EOF Then
        Flog.Writeline Espacios(Tabulador * 2) & "No se encontro el legajo " & Nro_Legajo
        InsertaError 1, 8
        HuboError = True
        Exit Sub
    Else
        Tercero = rs_Empleado!ternro
    End If
    
    Aux_Fecha_Desde = Fecha_Desde
    Aux_Dias = Total_Dias


    'Busco todos los Periodos Involucrados entre las Fechas
    StrSql = "SELECT * FROM vacacion "
    'StrSql = StrSql & " WHERE vacfecdesde <= " & ConvFecha(fecha_hasta)
    'StrSql = StrSql & " AND  vacfechasta >= " & ConvFecha(fecha_desde)
    'StrSql = StrSql & " ORDER BY vacnro"
    StrSql = StrSql & " ORDER BY vacfecdesde"
    OpenRecordset StrSql, rs_Periodos_Vac
    
    Do While Not rs_Periodos_Vac.EOF And Aux_Dias >= 0 'And Aux_Fecha_Desde < fecha_hasta
        'si le fecha en la que se va a generar los dias pedidos estan fuera del rengo de fechas del periodo
        ' no se procesan
        If Aux_Fecha_Desde >= rs_Periodos_Vac!vacfecdesde Then ' And Aux_Fecha_Desde <= rs_Periodos_Vac!vacfechasta Then
            diascoract = 0
            diastom = 0
            diascorant = 0
            diasdebe = 0
            diastot = 0
            diasyaped = 0
            diaspend = 0
            
            Flog.Writeline Espacios(Tabulador * 2) & "Periodo de Vacaciones:" & rs_Periodos_Vac!vacnro & " " & rs_Periodos_Vac!vacdesc
            
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
                'Call DiasPedidos(NroVac, nroTipvac, Aux_Fecha_Desde, Fecha_Hasta, Hasta, Tercero, diaspend, tothabiles, totNohabiles, totferiados)
                If diaspend > Aux_Dias Then
                    diaspend = Aux_Dias
                End If
                Call DiasPedidos(NroVac, nroTipvac, Aux_Fecha_Desde, Hasta, Tercero, diaspend, tothabiles, totNohabiles, totferiados)
                
                StrSql = "INSERT INTO vacdiasped (vdiapedhasta,vdiaspedferiados,vdiaspedhabiles,vdiaspednohabiles,ternro,vacnro,vdiapedcant,vdiapeddesde,vdiaspedestado) VALUES (" & _
                          ConvFecha(Hasta) & "," & totferiados & "," & tothabiles & "," & totNohabiles & "," & Tercero & "," & NroVac & "," & (totferiados + tothabiles + totNohabiles) & "," & ConvFecha(Aux_Fecha_Desde) & ",-1)"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Aux_Dias = Aux_Dias - tothabiles
                Aux_Fecha_Desde = Hasta + 1
            End If
        Else
            Flog.Writeline Espacios(Tabulador * 2) & "La fecha en la que se va a generar los dias pedidos estan fuera del rengo de fechas del periodo " & Aux_Fecha_Desde
        End If
            
        rs_Periodos_Vac.MoveNext
    Loop
    
    'If Aux_Fecha_Desde < Fecha_Hasta Then
    If Aux_Dias > 0 Then
        Flog.Writeline Espacios(Tabulador * 2) & "Quedaron " & Aux_Dias & " días sin generar "
    End If

Fin:

'Desactivo el manejador de Errores Local
On Error GoTo 0

Exit Sub

Manejador_De_Error:
    HuboError = True

    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strlinea
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin
End Sub

'NroVac,nroTipvac, Aux_Fecha_Desde, Fecha_Hasta, Hasta, Tercero, diaspend, tothabiles, totNohabiles, totferiados)
Private Sub DiasPedidos(ByVal NroVac As Long, ByVal TipoVac As Long, ByVal FechaInicial As Date, ByRef Fecha As Date, ByVal ternro As Long, ByVal cant As Long, ByRef cHabiles As Long, ByRef cNoHabiles As Long, ByRef cFeriados As Long)
'Calcula la fecha hasta a partir de la fecha desde, la cantidad de dias pedidos y el tipo
'de vacacion asociado a los dias correspòndientes, para el período

Dim i As Long
Dim j As Long
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
        Flog.Writeline Espacios(Tabulador * 2) & "No se encontro el tipo de Vacacion " & TipoVac
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
    
    Do While i <= cant
    
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
'            If DHabiles(Weekday(Fecha)) Or ((esFeriado) And ExcluyeFeriados) Then
'                i = i + 1
'            Else
'                cHabiles = cHabiles + 1
'            End If
        End If
        If i < cant Then
            Fecha = DateAdd("d", 1, Fecha)
        Else
            i = i + 1
        End If
    Loop
    
    Set objFeriado = Nothing

End Sub


Private Sub DiasPedidos_old(ByVal NroVac As Long, ByVal TipoVac As Long, ByVal FechaInicial As Date, ByVal FechaFinal As Date, ByRef Fecha As Date, ByVal ternro As Long, ByRef cant As Long, ByRef cHabiles As Long, ByRef cNoHabiles As Long, ByRef cFeriados As Long)
'calcula la cantidad de dias y la fecha hasta correspondiente al periodo
'de acuerdo al tipo de vacacion

Dim i As Long
Dim j As Long
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


Private Sub Calcular_Hasta(ByVal ternro As Long, ByVal TipoVac As Long, ByVal Fecha_Desde As Date, ByVal dias As Long, ByRef Fecha_Hasta As Date)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la Fecha hasta teniendo en cuenta la fecha desde, la cantidad de dias
'               y el tipo de vacaciones (Dias Corridos o Dias Habiles)
' Autor      : FGZ
' Fecha      : 21/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------



End Sub


Public Sub LineaModelo_231(ByVal strlinea As String)
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

Dim pos1 As Long
Dim pos2 As Long
    
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

On Error GoTo Manejador_De_Error

' El formato es: segun formato del tipo de datos TR_Datos_Bancarios
        
    'Proceso        long 2
    pos1 = 1
    pos2 = 2 + 1
    Registro.Proceso = Mid$(strlinea, pos1, pos2 - pos1)
    
    'Servicio       long 4
    pos1 = 3
    pos2 = 6 + 1
    Registro.Servicio = Mid$(strlinea, pos1, pos2 - pos1)
    
    'Sucursal       long 4
    pos1 = 7
    pos2 = 10 + 1
    Registro.Sucursal = Mid$(strlinea, pos1, pos2 - pos1)
    
    'Legajo         long 20
    pos1 = 11
    pos2 = 30 + 1
    Registro.Legajo = Mid$(strlinea, pos1, pos2 - pos1)
    
    'Moneda         long 1
    pos1 = 31
    pos2 = 31 + 1
    Registro.Moneda = Mid$(strlinea, pos1, pos2 - pos1)
    
    'Titularidad    long 2
    pos1 = 32
    pos2 = 33 + 1
    Registro.Titularidad = Mid$(strlinea, pos1, pos2 - pos1)
    
    'CBU long 22 -
        'Bloque 1
        '       Codigo entidad:  "011"  - long 3
                pos1 = 34
                pos2 = 36 + 1
                CBU_Bloque1.Codigo_Entidad = Mid$(strlinea, pos1, pos2 - pos1)
        '       Codigo Sucursal: "BCRA" - long 4
                pos1 = 37
                pos2 = 40 + 1
                CBU_Bloque1.Codigo_Sucursal = Mid$(strlinea, pos1, pos2 - pos1)
        '       Digito Verif. Bloque 1  - long 1
                pos1 = 41
                pos2 = 41 + 1
                CBU_Bloque1.Digito_Verificador = Mid$(strlinea, pos1, pos2 - pos1)
    
        'Bloque 2
        '       Tipo de Cuenta:         - long 1  (2 = CC, 3 = CA y 4 = CCE)
                pos1 = 42
                pos2 = 42 + 1
                CBU_Bloque2.Cuenta_Tipo = Mid$(strlinea, pos1, pos2 - pos1)
        '       Moneda de la cuenta:    - long 1  (0 = Pesos, 1 = Dolares y 3 = Lecop)
                pos1 = 43
                pos2 = 43 + 1
                CBU_Bloque2.Moneda = Mid$(strlinea, pos1, pos2 - pos1)
        '       Nro de la cuenta        - long 11
                pos1 = 45
                pos2 = 54 + 1
                CBU_Bloque2.Cuenta_Nro = Mid$(strlinea, pos1, pos2 - pos1)
        '       Digito Verif. Bloque 2  - long 1
                pos1 = 55
                pos2 = 55 + 1
                CBU_Bloque2.Digito_Verificador = Mid$(strlinea, pos1, pos2 - pos1)
    CBU.Bloque1 = CBU_Bloque1.Codigo_Entidad & CBU_Bloque1.Codigo_Sucursal & CBU_Bloque1.Digito_Verificador
    CBU.Bloque2 = CBU_Bloque2.Cuenta_Tipo & CBU_Bloque2.Moneda & CBU_Bloque2.Cuenta_Nro & CBU_Bloque2.Digito_Verificador
    Registro.CBU = CBU.Bloque1 & CBU.Bloque2
    
    'Cuenta_Electronica     long 19 - (nro de tarjeta de debito)
    pos1 = 56
    pos2 = 74 + 1
    Registro.Cuenta_Electronica = Mid$(strlinea, pos1, pos2 - pos1)
    'Tarjeta_1er_Titular    long 19 -
    pos1 = 75
    pos2 = 93 + 1
    Registro.Tarjeta_1er_Titular = Mid$(strlinea, pos1, pos2 - pos1)
    'Tarjeta_2do_Titular    long 19 -
    pos1 = 94
    pos2 = 112 + 1
    Registro.Tarjeta_2do_Titular = Mid$(strlinea, pos1, pos2 - pos1)
    'Doc_Tipo As String     long 2  -
    pos1 = 113
    pos2 = 114 + 1
    Registro.Doc_Tipo = Mid$(strlinea, pos1, pos2 - pos1)
    'Doc_Nro                long 11 -
    pos1 = 115
    pos2 = 125 + 1
    Registro.Doc_Nro = Mid$(strlinea, pos1, pos2 - pos1)
    'Filler                 long 5  -
    pos1 = 126
    pos2 = 130 + 1
    Registro.Filler = Mid$(strlinea, pos1, pos2 - pos1)
    
' ====================================================================
'   Validar los parametros Levantados
If EsNulo(Registro.Legajo) Then
    Flog.Writeline Espacios(Tabulador * 2) & "No se encontro el legajo " & Registro.Legajo
    InsertaError 1, 8
    HuboError = True
    Exit Sub
End If
'Que exista el legajo
StrSql = "SELECT * FROM empleado where empleg = " & Registro.Legajo
OpenRecordset StrSql, rs_Empleado
If rs_Empleado.EOF Then
    Flog.Writeline Espacios(Tabulador * 2) & "No se encontro el legajo " & Registro.Legajo
    InsertaError 1, 8
    HuboError = True
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
        HuboError = True
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
    HuboError = True
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
    HuboError = True
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
    
Fin:
'Cierro todo y libero
If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
If rs_CtaBancaria.State = adStateOpen Then rs_CtaBancaria.Close
If rs_FormaPago.State = adStateOpen Then rs_FormaPago.Close
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close

Set rs_CtaBancaria = Nothing
Set rs_Empleado = Nothing
Set rs_FormaPago = Nothing
Set rs_Confrep = Nothing
Exit Sub

Manejador_De_Error:
    HuboError = True

    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strlinea
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin


End Sub



Public Sub LineaModelo_233(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta Licencia segun formato
' Autor      : FGZ
'              El formato es: Legajo,
'                             tipo de licencia (Char de 25 buscar con descripcion de tipo de dia),
'                             fecha desde,
'                             fecha hasta,
'                             dia completo (si/no),
'                             hora desde,
'                             hora hasta,
'                             cant de horas.
' Fecha      : 22/10/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1 As Long
Dim pos2 As Long
    
Dim objHoras As New FechasHoras

Dim Tercero As Long
Dim NroLegajo As Long

Dim Licencia_Descripcion As String
Dim TDNro As Long
Dim DiaCompleto As Boolean
Dim FechaDesde As Date
Dim FechaHasta As Date
Dim HoraDesde As String
Dim HoraHasta As String
Dim CantidadHoras As Single
Dim emp_licnro As Long


Dim aux As String

Dim rs_Empleado As New ADODB.Recordset
Dim rs_Lic As New ADODB.Recordset
Dim rs_TipDia As New ADODB.Recordset

'Formato
' ------------------------------------------------------------------------------------------
'Campo              Columna Tipo de Dato    Formato         Ejemplo             Obligatorio
' ------------------------------------------------------------------------------------------
'Legajo                 1   Entero              9(6)        1
'Tipo de Licencia       2   Carácter            X(25)       Licencia por Examen
'Fecha Desde            3   Fecha           DD/MM/AAAA      01/01/2004
'Fecha Hasta            4   Fecha           DD/MM/AAAA      14/01/2004
'Día Completo           5   Logico              Si/No       Si
'Hora Desde             6   Carácter            X(5)        00:00
'Hora Hasta             7   Carácter            X(5)        23:59
'Cantidad de Horas      8   Decimal          9(15).999      8.5
' ------------------------------------------------------------------------------------------

On Error GoTo Manejador_De_Error

    'Nro de Legajo
    pos1 = 1
    pos2 = InStr(pos1, strlinea, Separador)
    If IsNumeric(Mid$(strlinea, pos1, pos2 - pos1)) Then
        NroLegajo = Mid$(strlinea, pos1, pos2 - pos1)
    Else
        InsertaError 1, 8
        HuboError = True
        Exit Sub
    End If
    
    'Tipo de Licencia
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Licencia_Descripcion = Mid(strlinea, pos1, pos2 - pos1)

    'Fecha desde
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    FechaDesde = Mid(strlinea, pos1, pos2 - pos1)
            
    'Fecha Hasta
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    FechaHasta = Mid(strlinea, pos1, pos2 - pos1)
                
    'Dia Completo
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    If pos2 > 0 Then
        aux = Mid(strlinea, pos1, pos2 - pos1)
        If UCase(aux) = "SI" Then
            DiaCompleto = True
        Else
            DiaCompleto = False
        End If
    Else
        pos2 = Len(strlinea)
        aux = CBool(Mid(strlinea, pos1, pos2 - pos1))
        If UCase(aux) = "S" Then
            DiaCompleto = True
        Else
            DiaCompleto = False
        End If
    End If
    
'    If Not DiaCompleto Then
        'Hora desde
        pos1 = pos2 + 1
        pos2 = InStr(pos1 + 1, strlinea, Separador)
        HoraDesde = Mid(strlinea, pos1, pos2 - pos1)
        
        'Hora hasta
        pos1 = pos2 + 1
        pos2 = InStr(pos1 + 1, strlinea, Separador)
        HoraHasta = Mid(strlinea, pos1, pos2 - pos1)
                
        'Cantidad de Horas
        pos1 = pos2 + 1
        pos2 = Len(strlinea) + 1
        CantidadHoras = CSng(Mid(strlinea, pos1, pos2 - pos1))
'    Else
'        HoraDesde = ""
'        HoraHasta = ""
'        CantidadHoras = 0
'    End If

' ====================================================================
'   Validar los parametros Levantados

'Que exista el legajo
StrSql = "SELECT * FROM empleado where empleg = " & NroLegajo
OpenRecordset StrSql, rs_Empleado
If rs_Empleado.EOF Then
    Flog.Writeline Espacios(Tabulador * 2) & "No se encontro el legajo " & NroLegajo
    InsertaError 1, 8
    HuboError = True
    Exit Sub
Else
    Tercero = rs_Empleado!ternro
End If

'Que exista el Tipo de Licencia
StrSql = "SELECT * FROM tipdia WHERE tddesc = '" & Licencia_Descripcion & "'"
OpenRecordset StrSql, rs_TipDia
If rs_TipDia.EOF Then
    Flog.Writeline Espacios(Tabulador * 2) & "No se encontro el Tipo de Licencia " & Licencia_Descripcion
    InsertaError 2, 84
    HuboError = True
    Exit Sub
Else
    TDNro = rs_TipDia!TDNro
End If

'Validar Horas
If Not DiaCompleto Then
    If Not objHoras.ValidarHora(HoraDesde) Then
        Flog.Writeline Espacios(Tabulador * 2) & "formato de Hora Desde incorrecto " & HoraDesde
        InsertaError 6, 98
        HuboError = True
        Exit Sub
    End If
    If Not objHoras.ValidarHora(HoraHasta) Then
        Flog.Writeline Espacios(Tabulador * 2) & "formato de Hora Hasta incorrecto " & HoraHasta
        InsertaError 7, 98
        HuboError = True
        Exit Sub
    End If
Else
    If Not objHoras.ValidarHora(HoraDesde) Then
        Flog.Writeline Espacios(Tabulador * 2) & "formato de Hora Desde incorrecto " & HoraDesde
        InsertaError 6, 98
        HuboError = True
        Exit Sub
    End If
    If Not objHoras.ValidarHora(HoraHasta) Then
        Flog.Writeline Espacios(Tabulador * 2) & "formato de Hora Hasta incorrecto " & HoraHasta
        InsertaError 7, 98
        HuboError = True
        Exit Sub
    End If
End If

'Valido la cantidad de Horas
If DiaCompleto Then
    If CantidadHoras = 0 Then
        CantidadHoras = rs_TipDia!tdcanthoras
    End If
    
    If CantidadHoras > rs_TipDia!tdcanthoras Then
        Flog.Writeline Espacios(Tabulador * 2) & "La cantidad de Horas " & CantidadHoras & " excede el maximo " & rs_TipDia!tdcanthoras
        CantidadHoras = rs_TipDia!tdcanthoras
    End If
End If


'=============================================================

'Busco si existe la Licencia
StrSql = "SELECT * FROM emp_lic " & _
         " WHERE (empleado = " & Tercero & " )" & _
         " AND elfechadesde <=" & ConvFecha(FechaHasta) & _
         " AND elfechahasta >= " & ConvFecha(FechaDesde)
OpenRecordset StrSql, rs_Lic

    If Not rs_Lic.EOF Then
            Flog.Writeline Espacios(Tabulador * 2) & "No se puede insertar la Licencia poqrue se superpone con otra ya existente"
            Flog.Writeline Espacios(Tabulador * 2) & "(" & rs_Lic!emp_licnro & ") desde " & rs_Lic!elfechadesde & " hasta " & rs_Lic!elfechahasta & " de tipo " & rs_Lic!TDNro
            InsertaError 1, 99
            HuboError = True
            Exit Sub
        Else
            'Inserto la Licencia
            StrSql = "INSERT INTO emp_lic ("
            StrSql = StrSql & "empleado,elfechadesde,elfechahasta,tdnro,eldiacompleto,eltipo"
            If Not DiaCompleto Then
                StrSql = StrSql & ",elhoradesde,elhorahasta"
            End If
            StrSql = StrSql & ",elcantdias,elcanthrs"
            StrSql = StrSql & ") VALUES ("
            StrSql = StrSql & Tercero
            StrSql = StrSql & "," & ConvFecha(FechaDesde)
            StrSql = StrSql & "," & ConvFecha(FechaHasta)
            StrSql = StrSql & "," & TDNro
            StrSql = StrSql & "," & CInt(DiaCompleto)
            
            If DiaCompleto Then
                StrSql = StrSql & ",1"
            Else
                StrSql = StrSql & ",2" ' Parcial Fija
                StrSql = StrSql & ",'" & HoraDesde & "'"
                StrSql = StrSql & ",'" & HoraHasta & "'"
            End If
            
            StrSql = StrSql & "," & (DateDiff("d", FechaDesde, FechaHasta) + 1)
            StrSql = StrSql & "," & CantidadHoras
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.Writeline Espacios(Tabulador * 2) & "Licencia insertada "
                
            emp_licnro = getLastIdentity(objConn, "emp_lic")
            
            'Genero la Justificacion
             StrSql = " INSERT INTO gti_justificacion ( jusanterior,juscodext,jusdesde,jusdiacompleto,jushasta,jussigla,jussistema,ternro,tjusnro,turnro,jushoradesde,jushorahasta,juseltipo,juselmaxhoras ) VALUES ("
             StrSql = StrSql & "-1"
             StrSql = StrSql & "," & emp_licnro
             StrSql = StrSql & "," & ConvFecha(FechaDesde)
             StrSql = StrSql & ",-1"
             StrSql = StrSql & "," & ConvFecha(FechaHasta)
             StrSql = StrSql & ",'LIC'"
             StrSql = StrSql & ",-1"
             StrSql = StrSql & "," & Tercero
             StrSql = StrSql & ",1"
             StrSql = StrSql & ",0"
             StrSql = StrSql & ",'" & HoraDesde & "'"
             StrSql = StrSql & ",'" & HoraHasta & "'"
             StrSql = StrSql & "," & TDNro
             StrSql = StrSql & "," & CantidadHoras
             StrSql = StrSql & ")"
             objConn.Execute StrSql, , adExecuteNoRecords
             Flog.Writeline Espacios(Tabulador * 2) & "Justificacion insertada "
        End If

Fin:
'Cierro todo y libero
If rs_Lic.State = adStateOpen Then rs_Lic.Close
If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
If rs_TipDia.State = adStateOpen Then rs_TipDia.Close

Set rs_Lic = Nothing
Set rs_Empleado = Nothing
Set rs_TipDia = Nothing
Exit Sub

Manejador_De_Error:
    HuboError = True

    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strlinea
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin

End Sub


Public Sub LineaModelo_235(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta en gti_rep_estad_accid_his.
' Autor      : FGZ
'              El formato es: estrnro,
'                             Periodo de GTI,
'                             cantidad de empleados,
'                             Dias trabajados,
'                             cant de horas extras.
' Fecha      : 03/11/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1 As Long
Dim pos2 As Long
    
Dim Estructura As Long
Dim Periodo_GTI As Long
Dim Cantidad_Empleados As Long
Dim Cantidad_Horas_Extras As Long
Dim Dias_Trabajados As Long


Dim rs_Gti_Rep_Estad_Accid_His As New ADODB.Recordset
Dim rs_Estructura  As New ADODB.Recordset
Dim rs_Periodo_GTI  As New ADODB.Recordset

On Error GoTo Manejador_De_Error

    'Estructura
    pos1 = 1
    pos2 = InStr(pos1, strlinea, Separador)
    If IsNumeric(Mid$(strlinea, pos1, pos2 - pos1)) Then
        Estructura = CLng(Mid$(strlinea, pos1, pos2 - pos1))
    Else
        InsertaError 1, 8
        HuboError = True
        Exit Sub
    End If
    
    'Periodo GTI
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Periodo_GTI = CLng(Mid(strlinea, pos1, pos2 - pos1))

    'Cantidad de Empleados
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Cantidad_Empleados = CInt(Mid(strlinea, pos1, pos2 - pos1))
        
    'Cantidad de dias
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Dias_Trabajados = CInt(Mid(strlinea, pos1, pos2 - pos1))
        
    'Cantidad de Horas
    pos1 = pos2 + 1
    pos2 = Len(strlinea) + 1
    Cantidad_Horas_Extras = CLng(Mid(strlinea, pos1, pos2 - pos1))

' ====================================================================
'   Validar los parametros Levantados
'Periodo
StrSql = "SELECT * FROM gti_per "
StrSql = StrSql & " WHERE pgtinro = " & Periodo_GTI
OpenRecordset StrSql, rs_Periodo_GTI

If rs_Periodo_GTI.EOF Then
    InsertaError 2, 50
    Flog.Writeline Espacios(Tabulador * 2) & "No existe ese periodo de GTI" & Periodo_GTI
    HuboError = True
    Exit Sub
End If

'Estructura
StrSql = "SELECT * FROM estructura "
StrSql = StrSql & " WHERE estrnro = " & Estructura
OpenRecordset StrSql, rs_Estructura

If rs_Estructura.EOF Then
    Flog.Writeline Espacios(Tabulador * 2) & "No existe esa estructura" & Estructura
    InsertaError 1, 100
    HuboError = True
    Exit Sub
End If


'=============================================================

'Busco si existe la Licencia
StrSql = "SELECT * FROM gti_rep_estad_accid_his "
StrSql = StrSql & " WHERE estrnro = " & Estructura
StrSql = StrSql & " AND pgtinro = " & Periodo_GTI
OpenRecordset StrSql, rs_Gti_Rep_Estad_Accid_His

    If Not rs_Gti_Rep_Estad_Accid_His.EOF Then
            Flog.Writeline Espacios(Tabulador * 2) & "No se puede insertar. Ya existenten datos para esa estructura en ese periodo de GTI"
            InsertaError 1, 103
            HuboError = True
            Exit Sub
        Else
            'Inserto la Licencia
            StrSql = "INSERT INTO gti_rep_estad_accid_his ("
            StrSql = StrSql & "estrnro,pgtinro,em,dt,hextras"
            StrSql = StrSql & ") VALUES ("
            StrSql = StrSql & Estructura
            StrSql = StrSql & "," & Periodo_GTI
            StrSql = StrSql & "," & Cantidad_Empleados
            StrSql = StrSql & "," & Dias_Trabajados
            StrSql = StrSql & "," & Cantidad_Horas_Extras
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.Writeline Espacios(Tabulador * 2) & "linea insertada "
        End If
Fin:
'Cierro todo y libero
If rs_Gti_Rep_Estad_Accid_His.State = adStateOpen Then rs_Gti_Rep_Estad_Accid_His.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_Periodo_GTI.State = adStateOpen Then rs_Periodo_GTI.Close

Set rs_Gti_Rep_Estad_Accid_His = Nothing
Set rs_Estructura = Nothing
Set rs_Periodo_GTI = Nothing

Exit Sub

Manejador_De_Error:
    HuboError = True

    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strlinea
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin

End Sub


Public Sub LineaModelo_236(ByVal strlinea As String)
Dim pos1 As Long
Dim pos2 As Long
    
'Configuración de conceptos para las novedades
Dim id_concepto_pera As Long
Dim id_concepto_manzana As Long
Dim id_concepto_carozo As Long
Dim concepto_pera As String
Dim concepto_manzana As String
Dim concepto_carozo As String

Dim id_th_bultos As Long '        AS INT INITIAL 51  . /* THora de BULTOS */.

Dim cant_bultos_txt As String
Dim monto_bultos_txt As String
Dim cant_bultos   As Single
Dim monto_bultos As Single
Dim primera_parte As String
Dim empaque As Long
Dim Legajo As Long
Dim Fecha_Desde As String
Dim Fecha_Hasta As String
Dim producto_txt As String
Dim producto As Long
Dim fecha_prod As Date

Dim rs_Empleado As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Sucursal As New ADODB.Recordset
Dim rs_Gti_Achdiario As New ADODB.Recordset

Dim fs1
Dim Flog1
Dim txtArchivoNov


On Error GoTo Manejador_De_Error

'Conccod de productos
concepto_pera = "pr050"
concepto_manzana = "pr120"
concepto_carozo = "pr140"

'Obtengo el código del concepto pera
StrSql = "SELECT * FROM concepto WHERE " & _
" conccod = '" & concepto_pera & "'"
OpenRecordset StrSql, rs_Gti_Achdiario
If Not rs_Gti_Achdiario.EOF Then
    id_concepto_pera = rs_Gti_Achdiario!concnro
Else
    Flog.Writeline Espacios(Tabulador * 1) & "No se encontro el Concepto para la Pera " & concepto_pera
    FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": No se encontro el Concepto " & concepto_pera
    InsertaError 2, 10
    HuboError = True
    Exit Sub
End If

'Obtengo el código del concepto manzana
StrSql = "SELECT * FROM concepto WHERE " & _
" conccod = '" & concepto_manzana & "'"
OpenRecordset StrSql, rs_Gti_Achdiario
If Not rs_Gti_Achdiario.EOF Then
    id_concepto_manzana = rs_Gti_Achdiario!concnro
Else
    Flog.Writeline Espacios(Tabulador * 1) & "No se encontro el Concepto para la Manzana " & concepto_manzana
    FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": No se encontro el Concepto " & concepto_manzana
    InsertaError 2, 10
    HuboError = True
    Exit Sub
End If

'Obtengo el código del concepto carozo
StrSql = "SELECT * FROM concepto WHERE " & _
" conccod = '" & concepto_carozo & "'"
OpenRecordset StrSql, rs_Gti_Achdiario
If Not rs_Gti_Achdiario.EOF Then
    id_concepto_carozo = rs_Gti_Achdiario!concnro
Else
    Flog.Writeline Espacios(Tabulador * 1) & "No se encontro el Concepto para el Carozo " & concepto_carozo
    FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": No se encontro el Concepto " & concepto_carozo
    InsertaError 2, 10
    HuboError = True
    Exit Sub
End If


'-----------------------------------------------------
'borrado de novedades

'Primero hago un backup de las novedades que voy a borrar
Set fs1 = CreateObject("Scripting.FileSystemObject")
txtArchivoNov = PathFLog & "novemp" & CStr(Format(Date, "yyyymmdd")) & Format(Time, "hhmm") & ".txt"
Set Flog1 = fs.CreateTextFile(txtArchivoNov, True)

StrSql = "SELECT *  FROM novemp WHERE " & _
" concnro = " & id_concepto_manzana & _
" OR concnro = " & id_concepto_pera & _
" OR concnro = " & id_concepto_carozo
OpenRecordset StrSql, rs_Gti_Achdiario
Do While Not rs_Gti_Achdiario.EOF
    Flog1.Write rs_Gti_Achdiario!concnro & "," & rs_Gti_Achdiario!tpanro & ","
    Flog1.Write rs_Gti_Achdiario!Empleado & "," & rs_Gti_Achdiario!nevalor & ","
    Flog1.Write rs_Gti_Achdiario!nevigencia & "," & rs_Gti_Achdiario!nedesde & ","
    Flog1.Write rs_Gti_Achdiario!nehasta & "," & rs_Gti_Achdiario!neretro & ","
    Flog1.Write rs_Gti_Achdiario!nepliqdesde & "," & rs_Gti_Achdiario!nepliqhasta & ","
    Flog1.Write rs_Gti_Achdiario!pronro & "," & rs_Gti_Achdiario!nenro & ","
    Flog1.Writeline
       
    rs_Gti_Achdiario.MoveNext
Loop

'Borro las novedades de peras, manzanas o carozo
StrSql = "DELETE FROM novemp WHERE " & _
" concnro = " & id_concepto_manzana & _
" OR concnro = " & id_concepto_pera & _
" OR concnro = " & id_concepto_carozo
objConn.Execute StrSql, , adExecuteNoRecords



'---------------------------------------------------------


'Levanto los datos que vienen en la linea
'strLinea1
empaque = CInt(Mid(strlinea, 1, 1))
Legajo = CLng(Mid(strlinea, 2, 6))
Fecha_Desde = Mid(strlinea, 8, 10)
Fecha_Hasta = Mid(strlinea, 18, 10)
producto_txt = Trim(Mid(strlinea, 28, 10))

If empaque = 1 Then
    empaque = 236
Else
    empaque = 237
End If

'strLinea2
'fecha_prod = CDate(Mid(strLinea, 1, 10))

'strLinea3
cant_bultos = CSng(Mid(strlinea, 38, 5))
'If Len(strLinea) > 0 Then
'    cant_bultos = cant_bultos + CInt(Mid(strLinea, Len(strLinea) - 1, 2) / 100)
'Else
'    cant_bultos = cant_bultos + CInt(Mid(strLinea, Len(strLinea) - 1, 2) / 100)
'End If

'strLinea4
monto_bultos = CSng(Mid(strlinea, 43, 7))
'If Len(strLinea) > 0 Then
'    monto_bultos = monto_bultos + CInt(Mid(strLinea, Len(strLinea) - 1, 2) / 100)
'Else
'    monto_bultos = monto_bultos + CInt(Mid(strLinea, Len(strLinea) - 1, 2) / 100)
'End If


' ====================================================================
' control de errores
    StrSql = "SELECT * FROM empleado where empleg = " & Legajo
    OpenRecordset StrSql, rs_Empleado
    If rs_Empleado.EOF Then
        Flog.Writeline "Empleado Inexistente: " & Legajo
        HuboError = True
        Exit Sub
    End If
    
   
    Select Case producto_txt
    Case "PERAS":
        producto = 1 'id-th-bultos = id-th-bultos-pera
    Case "MANZANAS":
        producto = 2 'id-th-bultos = id-th-bultos-manzana
    Case "DURAZNOS":
        producto = 3 'id-th-bultos = id-th-bultos-carozo
    Case "PELONES":
        producto = 3 'id-th-bultos = id-th-bultos-carozo
    Case "CIRUELAS":
        producto = 3 'id-th-bultos = id-th-bultos-carozo
    End Select

'   Hasta aquí el código anterior
'=============================================================

If empaque = 236 Then
    Select Case producto
    Case 1:      'PERAS
        'Busco la novedad. Si no existe la creo
        StrSql = "SELECT * FROM novemp WHERE " & _
         " concnro = " & id_concepto_pera & _
         " AND empleado = " & rs_Empleado!ternro & _
         " AND tpanro = 51"
        OpenRecordset StrSql, rs_Gti_Achdiario
        If rs_Gti_Achdiario.EOF Then
            StrSql = "INSERT INTO novemp (" & _
             "empleado,concnro,nevalor,tpanro" & _
             ") VALUES (" & rs_Empleado!ternro & _
             "," & id_concepto_pera & _
             "," & cant_bultos & _
             ", 51" & _
             " )"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    Case 2:          'MANZANAS
        'Busco la novedad. Si no existe la creo
        StrSql = "SELECT * FROM novemp WHERE " & _
         " concnro = " & id_concepto_manzana & _
         " AND empleado = " & rs_Empleado!ternro & _
         " AND tpanro = 163"
        OpenRecordset StrSql, rs_Gti_Achdiario
        If rs_Gti_Achdiario.EOF Then
            StrSql = "INSERT INTO novemp (" & _
             "empleado,concnro,nevalor,tpanro" & _
             ") VALUES (" & rs_Empleado!ternro & _
             "," & id_concepto_manzana & _
             "," & cant_bultos & _
             ", 163" & _
             " )"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        'Busco la novedad. Si no existe la creo
        StrSql = "SELECT * FROM novemp WHERE " & _
         " concnro = " & id_concepto_manzana & _
         " AND empleado = " & rs_Empleado!ternro & _
         " AND tpanro = 51"
        OpenRecordset StrSql, rs_Gti_Achdiario
        If rs_Gti_Achdiario.EOF Then
            StrSql = "INSERT INTO novemp (" & _
             "empleado,concnro,nevalor,tpanro" & _
             ") VALUES (" & rs_Empleado!ternro & _
             "," & id_concepto_manzana & _
             "," & cant_bultos & _
             ", 51" & _
             " )"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
    Case 3:          'CAROZO
        'Busco la novedad. Si no existe la creo
        StrSql = "SELECT * FROM novemp WHERE " & _
         " concnro = " & id_concepto_carozo & _
         " AND empleado = " & rs_Empleado!ternro & _
         " AND tpanro = 163"
        OpenRecordset StrSql, rs_Gti_Achdiario
        If rs_Gti_Achdiario.EOF Then
            StrSql = "INSERT INTO novemp (" & _
             "empleado,concnro,nevalor,tpanro" & _
             ") VALUES (" & rs_Empleado!ternro & _
             "," & id_concepto_carozo & _
             "," & cant_bultos & _
             ", 163" & _
             " )"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        'Busco la novedad. Si no existe la creo
        StrSql = "SELECT * FROM novemp WHERE " & _
         " concnro = " & id_concepto_carozo & _
         " AND empleado = " & rs_Empleado!ternro & _
         " AND tpanro = 51"
        OpenRecordset StrSql, rs_Gti_Achdiario
        If rs_Gti_Achdiario.EOF Then
            StrSql = "INSERT INTO novemp (" & _
             "empleado,concnro,nevalor,tpanro" & _
             ") VALUES (" & rs_Empleado!ternro & _
             "," & id_concepto_carozo & _
             "," & cant_bultos & _
             ", 51" & _
             " )"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End Select
End If

Fin:
Exit Sub

Manejador_De_Error:
    HuboError = True

    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strlinea
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin



End Sub


Public Sub LineaModelo_237(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: IMPORTACION DE Detalle de Cantidad de BULTOS  a  RH Pro
'              IDEA : importado un desglose de Acumulado Diario de GTI para un T.Hora espec¡fico.
'              Luego lee el archivo y crea los Desglose AD de GTI, siempre pisa (asumen un reg. x empleado x convinatoria)
'              Genera un log de error en el mismo TMP
'              configuraci¢n de CTTES para la IMPORTACION
' Autor      : FGZ
' Fecha      : 10/02/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1 As Long
Dim pos2 As Long
    
Dim id_producto_peras As Long '   AS INT INITIAL 1.
Dim id_producto_manzana As Long 'AS INT INITIAL 2.
Dim id_producto_carozo As Long 'AS INT INITIAL 3.  /* incluye: CIRUELA, DURAZNO Y PELONES */
Dim id_th_bultos As Long '        AS INT INITIAL 51  . /* THora de BULTOS */.

Dim cant_bultos_txt As String
Dim monto_bultos_txt As String
Dim cant_bultos   As Single
Dim monto_bultos As Single
Dim primera_parte As String
Dim empaque As Long
Dim Legajo As Long
Dim Fecha_Desde As String
Dim Fecha_Hasta As String
Dim producto_txt As String
Dim producto As Long
Dim fecha_prod As Date

Dim rs_Empleado As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Sucursal As New ADODB.Recordset
Dim rs_Gti_Achdiario As New ADODB.Recordset
Dim rs_gti_achdiario_Estr As New ADODB.Recordset

Dim achdnro As Long
Dim TipoEstr_Producto As Long
Dim TipoEstr_Sucursal As Long
Dim TipoEstr_Categoria As Long
Dim TipoEstr_RegimenHorario As Long

id_producto_peras = 1
id_producto_manzana = 2
id_producto_carozo = 3
id_th_bultos = 51

TipoEstr_Producto = 38
TipoEstr_Sucursal = 1
TipoEstr_Categoria = 3
TipoEstr_RegimenHorario = 21

On Error GoTo Manejador_De_Error

'Levanto los datos que vienen en la linea
'strLinea1
empaque = CInt(Mid(strlinea, 1, 1))
Legajo = CLng(Mid(strlinea, 2, 6))
Fecha_Desde = Mid(strlinea, 8, 10)
Fecha_Hasta = Mid(strlinea, 18, 10)
producto_txt = Trim(Mid(strlinea, 28, 10))

If empaque = 1 Then
    empaque = 206
Else
    empaque = 207
End If

'strLinea2
'fecha_prod = CDate(Mid(strLinea, 1, 10))

'strLinea3
cant_bultos = CSng(Mid(strlinea, 38, 5))
'If Len(strLinea) > 0 Then
'    cant_bultos = cant_bultos + CInt(Mid(strLinea, Len(strLinea) - 1, 2) / 100)
'Else
'    cant_bultos = cant_bultos + CInt(Mid(strLinea, Len(strLinea) - 1, 2) / 100)
'End If

'strLinea4
'monto_bultos = CSng(Mid(strLinea, 44, 7))


' ====================================================================
' control de errores
    StrSql = "SELECT * FROM empleado where empleg = " & Legajo
    OpenRecordset StrSql, rs_Empleado
    If rs_Empleado.EOF Then
        Flog.Writeline "Empleado Inexistente: " & Legajo
        HuboError = True
        Exit Sub
    End If
    
   
    Select Case producto_txt
    Case "PERAS":
        producto = 1 'id-th-bultos = id-th-bultos-pera
    Case "MANZANAS":
        producto = 2 'id-th-bultos = id-th-bultos-manzana
    Case "DURAZNOS":
        producto = 3 'id-th-bultos = id-th-bultos-carozo
    Case "PELONES":
        producto = 3 'id-th-bultos = id-th-bultos-carozo
    Case "CIRUELAS":
        producto = 3 'id-th-bultos = id-th-bultos-carozo
    End Select

    'Busco la Sucursal
    StrSql = " SELECT estrnro FROM his_estructura " & _
             " WHERE ternro = " & rs_Empleado!ternro & " AND " & _
             " tenro = 1 AND estrnro = " & empaque & " AND " & _
             " (htetdesde <= " & ConvFecha(Fecha_Hasta) & ") AND " & _
             " ((" & ConvFecha(Fecha_Hasta) & " <= htethasta) or (htethasta is null))"
    OpenRecordset StrSql, rs_Estructura

    If Not rs_Estructura.EOF Then
        StrSql = " SELECT * FROM sucursal " & _
                 " WHERE estrnro =" & rs_Estructura!estrnro
        OpenRecordset StrSql, rs_Sucursal
        
        If rs_Sucursal.EOF Then
            Flog.Writeline "Empaque Inexistente: " & empaque
            HuboError = True
            Exit Sub
        Else
        
        End If
    Else
        Flog.Writeline "Empaque Inexistente: " & empaque
        HuboError = True
        Exit Sub
    End If


'=============================================================
' Inserto
StrSql = "SELECT * FROM gti_achdiario WHERE " & _
         " ternro = " & rs_Empleado!ternro & _
         " AND thnro = " & id_th_bultos & _
         " AND achdfecha = " & ConvFecha(fecha_prod)
OpenRecordset StrSql, rs_Gti_Achdiario

If rs_Gti_Achdiario.EOF Then
    StrSql = "INSERT INTO gti_achdiario (" & _
             "ternro,thnro,achdfecha,achdcanthoras " & _
             ") VALUES (" & rs_Empleado!ternro & _
             "," & id_th_bultos & _
             "," & ConvFecha(fecha_prod) & _
             "," & cant_bultos & _
             " )"
    objConn.Execute StrSql, , adExecuteNoRecords
    achdnro = getLastIdentity(objConn, "gti_achdiario")
    
    'Tengo que insertar 3 registros en gti_achdiario_estr(uno por cada estructura: Producto, Sucursal)
    
    'Estructura producto
    StrSql = "INSERT INTO gti_achdiario_estr "
    StrSql = StrSql & " ( achdnro, achdfecha, tenro, estrnro ) "
    StrSql = StrSql & " VALUES ("
    StrSql = StrSql & achdnro
    StrSql = StrSql & "," & ConvFecha(fecha_prod)
    StrSql = StrSql & "," & TipoEstr_Producto
    StrSql = StrSql & "," & producto
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Estructura Sucursal
    StrSql = "INSERT INTO gti_achdiario_estr "
    StrSql = StrSql & " ( achdnro, achdfecha, tenro, estrnro ) "
    StrSql = StrSql & " VALUES ("
    StrSql = StrSql & achdnro
    StrSql = StrSql & "," & ConvFecha(fecha_prod)
    StrSql = StrSql & "," & TipoEstr_Sucursal
    StrSql = StrSql & "," & empaque
    objConn.Execute StrSql, , adExecuteNoRecords
Else
    'reviso que existan las estructuras
    'Producto
    StrSql = "SELECT * FROM gti_achdiario_estr WHERE "
    StrSql = StrSql & " achnro = " & rs_Gti_Achdiario!achdnro
    StrSql = StrSql & " AND achdfecha = " & ConvFecha(fecha_prod)
    StrSql = StrSql & " AND tenro = " & TipoEstr_Producto
    StrSql = StrSql & " AND estrnro = " & producto
    OpenRecordset StrSql, rs_gti_achdiario_Estr
    
    If rs_gti_achdiario_Estr.EOF Then
        StrSql = "INSERT INTO gti_achdiario_estr "
        StrSql = StrSql & " ( achdnro, achdfecha, tenro, estrnro ) "
        StrSql = StrSql & " VALUES ("
        StrSql = StrSql & rs_Gti_Achdiario!achdnro
        StrSql = StrSql & "," & ConvFecha(fecha_prod)
        StrSql = StrSql & "," & TipoEstr_Producto
        StrSql = StrSql & "," & producto
        objConn.Execute StrSql, , adExecuteNoRecords
    End If

    'Sucursal
    StrSql = "SELECT * FROM gti_achdiario_estr WHERE "
    StrSql = StrSql & " achnro = " & rs_Gti_Achdiario!achdnro
    StrSql = StrSql & " AND achdfecha = " & ConvFecha(fecha_prod)
    StrSql = StrSql & " AND tenro = " & TipoEstr_Sucursal
    StrSql = StrSql & " AND estrnro = " & empaque
    OpenRecordset StrSql, rs_gti_achdiario_Estr
    
    If rs_gti_achdiario_Estr.EOF Then
        StrSql = "INSERT INTO gti_achdiario_estr "
        StrSql = StrSql & " ( achdnro, achdfecha, tenro, estrnro ) "
        StrSql = StrSql & " VALUES ("
        StrSql = StrSql & rs_Gti_Achdiario!achdnro
        StrSql = StrSql & "," & ConvFecha(fecha_prod)
        StrSql = StrSql & "," & TipoEstr_Sucursal
        StrSql = StrSql & "," & empaque
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

Fin:
Exit Sub

Manejador_De_Error:
    HuboError = True

    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strlinea
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin


End Sub

Public Sub LineaModelo_243(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Interface de cuentas Corrrientes para los empleados.
' Autor      : FGZ
' Fecha      : 02/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1 As Long
Dim pos2 As Long
    
Dim Legajo As Long
Dim Tipo_Doc As String
Dim Nro_Doc As String
Dim Nro_Cuenta As String
Dim Forma_Pago As Long
Dim Banco As Long
Dim Estado As Boolean
Dim Sucursal As String
Dim CtaAcred As String
    
Dim Tercero As Long
Dim FechaDesde As String
Dim FechaHasta As String
Dim Tercero_Banco As Long
Dim Cbnro As Long
Dim CtabSuc As String

Dim rs_Empleado As New ADODB.Recordset
Dim rs_CtaBancaria As New ADODB.Recordset
Dim rs_FormaPago As New ADODB.Recordset
Dim rs_Doc As New ADODB.Recordset
Dim rs_Banco As New ADODB.Recordset

On Error GoTo Manejador_De_Error

' El formato es:
'   Legajo;
'   Tipo_Doc;
'   Nro_Doc;
'   Nro_Cuenta;
'   Forma_Pago;
'   Banco(estrnro);

'   [ Sucursal;
'   Cta de Acreditacion ]
        
    'Legajo
    pos1 = 1
    pos2 = InStr(pos1, strlinea, Separador)
    If pos1 < pos2 Then
        Legajo = Mid$(strlinea, pos1, pos2 - pos1)
    End If
    
    'Tipo de Doc
    pos1 = pos2 + 1
    pos2 = InStr(pos1, strlinea, Separador)
    Tipo_Doc = Mid$(strlinea, pos1, pos2 - pos1)
    
    'Nro_DOC
    pos1 = pos2 + 1
    If EsNulo(Tipo_Doc) Then
        pos2 = InStr(pos1, strlinea, Separador)
    Else
        pos2 = InStr(pos1 + 1, strlinea, Separador)
    End If
    Nro_Doc = Mid$(strlinea, pos1, pos2 - pos1)
    
    'Nro_cuenta
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Nro_Cuenta = Mid$(strlinea, pos1, pos2 - pos1)
    
    'Forma de Pago
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Forma_Pago = Mid$(strlinea, pos1, pos2 - pos1)
    
    'Banco (estrnro)
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    If pos2 = 0 Then
        pos2 = Len(strlinea)
        Banco = Mid$(strlinea, pos1, pos2)
        Sucursal = ""
        CtaAcred = ""
    Else
        Banco = Mid$(strlinea, pos1, pos2 - pos1)
    
        'Sucursal
        pos1 = pos2 + 1
        pos2 = InStr(pos1 + 1, strlinea, Separador)
        Sucursal = Mid$(strlinea, pos1, pos2 - pos1)
    
        'Cuenta de Acreditacion
        pos1 = pos2 + 1
        pos2 = Len(strlinea)
        CtaAcred = Mid$(strlinea, pos1, pos2)
    End If
    
' ====================================================================
'   Validar los parametros Levantados

'El estado es siempre activo
Estado = True

'Puede venir el legajo o el tipo y nro de documento
'Que exista el legajo
If Not Legajo = 0 Then
    StrSql = "SELECT * FROM empleado where empleg = " & Legajo
    OpenRecordset StrSql, rs_Empleado
    If rs_Empleado.EOF Then
        Flog.Writeline Espacios(Tabulador * 2) & "No se encontro el legajo " & Legajo
        InsertaError 1, 8
        HuboError = True
        Exit Sub
    Else
        Tercero = rs_Empleado!ternro
    End If
Else
    'Busco un empleado que tenga el mismo tipo y nro de doc
    StrSql = " SELECT ter_doc.nrodoc, tercero.ternro FROM tercero "
    StrSql = StrSql & " INNER JOIN ter_doc ON (tercero.ternro = ter_doc.ternro) "
    StrSql = StrSql & " INNER JOIN tipodocu ON (tipodocu.tidnro = ter_doc.tidnro AND tipodocu.tidsigla = '" & Tipo_Doc & "') "
    StrSql = StrSql & " WHERE nrodoc = '" & Nro_Doc & "'"
    If rs_Doc.State = adStateOpen Then rs_Doc.Close
    OpenRecordset StrSql, rs_Doc
    
    If Not rs_Doc.EOF Then
        Tercero = rs_Doc!ternro
    Else
        Flog.Writeline Espacios(Tabulador * 2) & "No se encontro el documento " & Tipo_Doc & " - " & Nro_Doc
        InsertaError 2, 31
        HuboError = True
        Exit Sub
    End If
End If

'Busco la forma de pago
StrSql = "SELECT * FROM formapago WHERE fpagbanc = -1 "
StrSql = StrSql & " AND fpagnro = " & Forma_Pago
If rs_FormaPago.State = adStateOpen Then rs_FormaPago.Close
OpenRecordset StrSql, rs_FormaPago
If rs_FormaPago.EOF Then
    Flog.Writeline Espacios(Tabulador * 2) & "No se encontro la forma de Pago " & Forma_Pago
    InsertaError 5, 104
    HuboError = True
    Exit Sub
Else
    Forma_Pago = rs_FormaPago!fpagnro
End If


'Busco el Banco
StrSql = "SELECT * FROM Banco "
StrSql = StrSql & " WHERE estrnro = " & Banco
If rs_Banco.State = adStateOpen Then rs_Banco.Close
OpenRecordset StrSql, rs_Banco
If rs_Banco.EOF Then
    Flog.Writeline Espacios(Tabulador * 2) & "No se encontro el Banco " & Banco
    InsertaError 6, 60
    HuboError = True
    Exit Sub
Else
    Tercero_Banco = rs_Banco!ternro
End If

'Sucursal
CtabSuc = Left(Sucursal, 10)

'Ctabacred
StrSql = "SELECT * FROM ctabancaria"
StrSql = StrSql & " WHERE ctabnro = '" & CtaAcred & "'"
If rs_CtaBancaria.State = adStateOpen Then rs_CtaBancaria.Close
OpenRecordset StrSql, rs_CtaBancaria
If Not rs_CtaBancaria.EOF Then
    Cbnro = rs_CtaBancaria!Cbnro
Else
    Flog.Writeline Espacios(Tabulador * 2) & "No se encontro el cbnro. Valor por default = 0 " & CtaAcred
    Cbnro = 0
End If
If rs_CtaBancaria.State = adStateOpen Then rs_CtaBancaria.Close

'=============================================================

'Busco si existe una cuenta para ese legajo para el mismo banco del mismo tipo de cuenta y activa
StrSql = "SELECT * FROM ctabancaria"
StrSql = StrSql & " WHERE ctabancaria.ternro =" & Tercero
StrSql = StrSql & " AND ctabestado = -1 "
StrSql = StrSql & " AND banco =" & Tercero_Banco
StrSql = StrSql & " AND fpagnro =" & Forma_Pago
If rs_CtaBancaria.State = adStateOpen Then rs_CtaBancaria.Close
OpenRecordset StrSql, rs_CtaBancaria
If Not rs_CtaBancaria.EOF Then
    StrSql = "UPDATE ctabancaria SET "
    StrSql = StrSql & " ctabsuc = '" & CtabSuc & "'"
    StrSql = StrSql & ", ctabacred = " & Cbnro
    StrSql = StrSql & ", ctabporc = 100 "
    StrSql = StrSql & " WHERE ctabancaria.ternro =" & Tercero
    StrSql = StrSql & " AND ctabestado = -1 "
    StrSql = StrSql & " AND banco =" & Tercero_Banco
    StrSql = StrSql & " AND fpagnro =" & Forma_Pago
    objConn.Execute StrSql, , adExecuteNoRecords

    Flog.Writeline Espacios(Tabulador * 2) & "Ya existe un cuenta del mismo tipo para el mismo banco activa. Se actualizará la cuenta"
Else
    StrSql = "INSERT INTO ctabancaria ("
    StrSql = StrSql & " ternro,fpagnro,banco,ctabestado, ctabnro, ctabsuc, ctabacred,ctabporc"
    StrSql = StrSql & ") VALUES (" & Tercero
    StrSql = StrSql & "," & Forma_Pago
    StrSql = StrSql & "," & Tercero_Banco
    StrSql = StrSql & ",-1"
    StrSql = StrSql & ",'" & Nro_Cuenta & "'"
    StrSql = StrSql & ",'" & CtabSuc & "'"
    StrSql = StrSql & "," & Cbnro
    StrSql = StrSql & ",100"
    StrSql = StrSql & " )"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline Espacios(Tabulador * 2) & "Cuenta Creada"
End If
    
Fin:
'Cierro todo y libero
If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
If rs_CtaBancaria.State = adStateOpen Then rs_CtaBancaria.Close
If rs_FormaPago.State = adStateOpen Then rs_FormaPago.Close
If rs_Doc.State = adStateOpen Then rs_Doc.Close
If rs_Banco.State = adStateOpen Then rs_Banco.Close

Set rs_CtaBancaria = Nothing
Set rs_Empleado = Nothing
Set rs_FormaPago = Nothing
Set rs_Doc = Nothing
Set rs_Banco = Nothing
Exit Sub

Manejador_De_Error:
    HuboError = True

    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strlinea
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin


End Sub

Public Sub LineaModelo_245(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta Novedad de ajuste segun formato
' Autor      : FGZ
'              El formato es:
'                   Formato 1
'                       Legajo; conccod; monto
'                   ó
'                   Formato 2.1
'                       Legajo; conccod; monto; FechaDesde; FechaHasta
'                   Formato 2.2
'                       Legajo; conccod; monto; FechaDesde
'                   ó
'                   Formato 3
'                       Legajo; conccod; monto; MarcaRetroactividad;PeriodoDesde(pliqnro); PeriodoHasta(pliqnro)
' Fecha      : 27/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1 As Long
Dim pos2 As Long
    
Dim Tercero As Long
Dim NroLegajo As Long

Dim concnro As Long
Dim Conccod As Long

Dim Monto As Single
Dim FechaDesde As String
Dim FechaHasta As String

Dim PeriodoDesde As Long
Dim PeriodoHasta As Long
Dim TieneVigencia As Boolean
Dim EsRetroactivo As Boolean

Dim aux As String

Dim rs_Empleado As New ADODB.Recordset
Dim rs_Concepto As New ADODB.Recordset
Dim rs_NovAju As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset


' El formato es:
' Formato 1
' Legajo; conccod; monto
' ó
' Formato 2.1
' Legajo; conccod; monto; FechaDesde; FechaHasta
' ó
' Formato 2.2
' Legajo; conccod; monto; FechaDesde
' ó
' Formato 3
' Legajo; conccod; monto; MarcaRetroactividad;PeriodoDesde(pliqnro); PeriodoHasta(pliqnro)
' ó
' Formato 4
' Legajo; conccod; monto; FechaDesde; FechaHasta; MarcaRetroactividad;PeriodoDesde(pliqnro); PeriodoHasta(pliqnro)
    
    On Error GoTo Manejador_De_Error
    
    TieneVigencia = False
    EsRetroactivo = False

    'Nro de Legajo
    pos1 = 1
    pos2 = InStr(pos1, strlinea, Separador)
    If IsNumeric(Mid$(strlinea, pos1, pos2 - pos1)) Then
        NroLegajo = Mid$(strlinea, pos1, pos2 - pos1)
    Else
        Flog.Writeline Espacios(Tabulador * 1) & "El legajo no es numerico "
        FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": El legajo no es numerico"
        InsertaError 1, 8
        HuboError = True
        Exit Sub
    End If
    
    'Concepto
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Conccod = Mid(strlinea, pos1, pos2 - pos1)

    'Monto
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    If pos2 = 0 Then
        pos2 = Len(strlinea)
        Monto = Mid(strlinea, pos1, pos2)
        Monto = CSng(Replace(CStr(Monto), SeparadorDecimal, "."))
    Else
        Monto = Mid(strlinea, pos1, pos2 - pos1)
        Monto = CSng(Replace(CStr(Monto), SeparadorDecimal, "."))
        
        'Puede veniar Fecha Desde; fecha Hasta ó Retroactivo, Periodo desde , Periodo Hasta
        pos1 = pos2 + 1
        pos2 = InStr(pos1 + 1, strlinea, Separador)
        If pos2 > 0 Then
            aux = Mid(strlinea, pos1, pos2 - pos1)
            If IsDate(aux) Then
                TieneVigencia = True
                'Fecha desde
                FechaDesde = Mid(strlinea, pos1, pos2 - pos1)
            
                'Fecha Hasta
                pos1 = pos2 + 1
                pos2 = InStr(pos1, strlinea, Separador)
                If pos2 > 0 Then
                    FechaHasta = Mid(strlinea, pos1, pos2 - pos1)
                    If IsDate(FechaHasta) Then
                        FechaHasta = CDate(FechaHasta)
                    Else
                        If Not EsNulo(FechaHasta) Then
                            Flog.Writeline Espacios(Tabulador * 1) & "Fecha no valida "
                            FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": La fecha no es valida "
                            InsertaError 1, 4
                            HuboError = True
                            Exit Sub
                        End If
                    End If
                    'Marca de Retroactividad
                    pos1 = pos2 + 1
                    pos2 = InStr(pos1 + 1, strlinea, Separador)
                    aux = Mid(strlinea, pos1, pos2 - pos1)
                    If UCase(aux) = "SI" Then
                        EsRetroactivo = True
                    Else
                        EsRetroactivo = False
                    End If
                
                    'Periodo desde
                    pos1 = pos2 + 1
                    pos2 = InStr(pos1 + 1, strlinea, Separador)
                    PeriodoDesde = Mid(strlinea, pos1, pos2 - pos1)
                    
                    'Periodo hasta
                    pos1 = pos2 + 1
                    pos2 = Len(strlinea)
                    PeriodoHasta = Mid(strlinea, pos1, pos2)
                Else
                    pos2 = Len(strlinea)
                    FechaHasta = Mid(strlinea, pos1, pos2)
                
                    TieneVigencia = True
                End If
            Else
                If UCase(aux) = "SI" Then
                    EsRetroactivo = True
                Else
                    EsRetroactivo = False
                End If
                
                'Periodo desde
                pos1 = pos2 + 1
                pos2 = InStr(pos1 + 1, strlinea, Separador)
                PeriodoDesde = Mid(strlinea, pos1, pos2 - pos1)
                
                'Periodo hasta
                pos1 = pos2 + 1
                pos2 = Len(strlinea)
                PeriodoHasta = Mid(strlinea, pos1, pos2)
            End If
        Else
            'Viene Vigencia con fecha desde y sin fecha hasta
            pos2 = Len(strlinea)
            FechaDesde = Mid(strlinea, pos1, pos2)
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
    Flog.Writeline Espacios(Tabulador * 1) & "No se encontro el legajo " & NroLegajo
    FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": No se encontro el legajo " & NroLegajo
    InsertaError 1, 8
    HuboError = True
    Exit Sub
Else
    Tercero = rs_Empleado!ternro
End If

'Que exista el concepto
StrSql = "SELECT * FROM concepto WHERE conccod = " & Conccod
StrSql = StrSql & " OR conccod = '" & Conccod & "'"
OpenRecordset StrSql, rs_Concepto
If rs_Concepto.EOF Then
    Flog.Writeline Espacios(Tabulador * 1) & "No se encontro el Concepto " & Conccod
    FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": No se encontro el Concepto " & Conccod
    InsertaError 2, 10
    HuboError = True
    Exit Sub
Else
    concnro = rs_Concepto!concnro
End If

If EsRetroactivo Then
    'Chequeo que los periodos sean validos
    'Chequeo Periodo Desde
    StrSql = "SELECT * FROM periodo WHERE pliqnro = " & PeriodoDesde
    If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
    OpenRecordset StrSql, rs_Periodo
    
    If rs_Periodo.EOF Then
        Flog.Writeline Espacios(Tabulador * 1) & "Periodo Desde Invalido " & PeriodoDesde
        FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": Periodo Desde Invalido " & PeriodoDesde
        InsertaError 6, 36
        HuboError = True
        Exit Sub
    End If
    
    'Chequeo Periodo Hasta
    StrSql = "SELECT * FROM periodo WHERE pliqnro = " & PeriodoHasta
    If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
    OpenRecordset StrSql, rs_Periodo
    
    If rs_Periodo.EOF Then
        Flog.Writeline Espacios(Tabulador * 1) & "Periodo Hasta Invalido " & PeriodoHasta
        FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": Periodo Hasta Invalido " & PeriodoHasta
        InsertaError 7, 36
        HuboError = True
        Exit Sub
    End If
End If

'=============================================================
'Busco si existe la Novedad
If Not TieneVigencia Then
    StrSql = "SELECT * FROM novaju WHERE "
    StrSql = StrSql & " concnro = " & concnro
    StrSql = StrSql & " AND empleado = " & Tercero
    StrSql = StrSql & " AND (navigencia = -1 OR navigencia = 0) "
    StrSql = StrSql & " ORDER BY navigencia "
    If rs_NovAju.State = adStateOpen Then rs_NovAju.Close
    OpenRecordset StrSql, rs_NovAju

    If Not rs_NovAju.EOF Then
        If CBool(rs_NovAju!navigencia) Then
            Flog.Writeline Espacios(Tabulador * 1) & "No se puede insertar la novedad poqrue ya existe una con Vigencia"
            FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": No se puede insertar la novedad poqrue ya existe una con Vigencia"
            InsertaError 1, 94
            HuboError = True
            Exit Sub
        Else
            'Existe una novedad pero sin vigencia ==> Actualizo
            If PisaNovedad Then 'Actualizo la Novedad
                If Not EsRetroactivo Then
                    StrSql = "UPDATE novaju SET navalor = " & Monto & _
                             " WHERE concnro = " & concnro & _
                             " AND empleado = " & Tercero
                Else
                    StrSql = "UPDATE novaju SET navalor = " & Monto & _
                             " , napliqdesde =  " & PeriodoDesde & _
                             " , napliqhasta =  " & PeriodoHasta & _
                             " WHERE concnro = " & concnro & _
                             " AND empleado = " & Tercero
                End If
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.Writeline Espacios(Tabulador * 1) & "Novedad Actualizada "
            Else
                Flog.Writeline Espacios(Tabulador * 1) & "No se insertó la novedad. Ya existe y no se pisa."
            End If
        End If
    Else
        'Inserto
        If Not EsRetroactivo Then
            StrSql = "INSERT INTO novaju (" & _
                     "empleado,concnro,navalor,navigencia" & _
                     ") VALUES (" & Tercero & _
                     "," & concnro & _
                     "," & Monto & _
                     ",0" & _
                     " )"
        Else
            StrSql = "INSERT INTO novaju (" & _
                     "empleado,concnro,navalor,navigencia,napliqdesde,napliqhasta " & _
                     ") VALUES (" & Tercero & _
                     "," & concnro & _
                     "," & Monto & _
                     "," & CInt(TieneVigencia) & _
                     "," & PeriodoDesde & _
                     "," & PeriodoHasta & _
                     " )"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.Writeline Espacios(Tabulador * 1) & "Novedad insertada "
    End If
Else 'Tiene Vigencia
    'Reviso que no se pise
    StrSql = "SELECT * FROM novaju WHERE "
    StrSql = StrSql & " concnro = " & concnro
    StrSql = StrSql & " AND empleado = " & Tercero
    StrSql = StrSql & " AND (navigencia = 0 "
    StrSql = StrSql & " OR (navigencia = -1 "
    If Not EsNulo(FechaHasta) Then
        StrSql = StrSql & " AND (nadesde <= " & ConvFecha(FechaHasta)
        StrSql = StrSql & " AND nahasta >= " & ConvFecha(FechaDesde) & ")"
        StrSql = StrSql & " OR  (nadesde <= " & ConvFecha(FechaHasta)
        StrSql = StrSql & " AND nahasta is null )))"
    Else
        StrSql = StrSql & " AND nahasta is null OR nahasta >= " & ConvFecha(FechaDesde) & "))"
    End If
    If rs_NovAju.State = adStateOpen Then rs_NovAju.Close
    OpenRecordset StrSql, rs_NovAju

    If Not rs_NovAju.EOF Then
        Flog.Writeline Espacios(Tabulador * 1) & "No se puede insertar la novedad, las vigencias se superponen"
        FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": No se puede insertar la novedad, las vigencias se superponen"
        InsertaError 1, 95
        HuboError = True
        Exit Sub
    Else
        If Not EsRetroactivo Then
            StrSql = "INSERT INTO novaju ("
            StrSql = StrSql & "empleado,concnro,navalor,navigencia,nadesde"
            If Not EsNulo(FechaHasta) Then
                StrSql = StrSql & ",nahasta"
            End If
            StrSql = StrSql & ") VALUES (" & Tercero
            StrSql = StrSql & "," & concnro
            StrSql = StrSql & "," & Monto
            StrSql = StrSql & ",-1"
            StrSql = StrSql & "," & ConvFecha(FechaDesde)
            If Not EsNulo(FechaHasta) Then
                StrSql = StrSql & "," & ConvFecha(FechaHasta)
            End If
            StrSql = StrSql & " )"
        Else
            StrSql = "INSERT INTO novaju ("
            StrSql = StrSql & "empleado,concnro,navalor,navigencia,nadesde"
            If Not EsNulo(FechaHasta) Then
                StrSql = StrSql & ",nahasta"
            End If
            StrSql = StrSql & ",napliqdesde,napliqhasta"
            StrSql = StrSql & ") VALUES (" & Tercero
            StrSql = StrSql & "," & concnro
            StrSql = StrSql & "," & Monto
            StrSql = StrSql & ",-1"
            StrSql = StrSql & "," & ConvFecha(FechaDesde)
            If Not EsNulo(FechaHasta) Then
                StrSql = StrSql & "," & ConvFecha(FechaHasta)
            End If
            StrSql = StrSql & "," & PeriodoDesde
            StrSql = StrSql & "," & PeriodoHasta
            StrSql = StrSql & " )"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.Writeline Espacios(Tabulador * 1) & "Novedad insertada "
    End If
End If

Fin:
'Cierro todo y libero
If rs_NovAju.State = adStateOpen Then rs_NovAju.Close
If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close

Set rs_NovAju = Nothing
Set rs_Empleado = Nothing
Set rs_Concepto = Nothing
Set rs_Periodo = Nothing
Exit Sub

Manejador_De_Error:
    HuboError = True

    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strlinea
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin
End Sub

Public Sub LineaModelo_246(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta en valgrilla.
'              El formato es:
'                   Nro. Escala; Coord1; Coord2; Coord3; Coord4; Coord5; Orden; Valor
' Autor      : Scarpa D.
' Fecha      : 21/01/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1 As Long
Dim pos2 As Long
Dim i As Long

Dim NroEscala As Long
Dim Coord1 As String
Dim Coord2 As String
Dim Coord3 As String
Dim Coord4 As String
Dim Coord5 As String
Dim Orden As Long
Dim Valor As Single

Dim ValorCorrecto As Boolean

Dim rs_Escala As New ADODB.Recordset
Dim rs_ValEscala As New ADODB.Recordset

On Error GoTo Manejador_De_Error
' El formato es:
' Nro. Escala; Coord1; Coord2; Coord3; Coord4; Coord5; Orden; Valor

    'Nro de Escala
    pos1 = 1
    pos2 = InStr(pos1, strlinea, Separador)
    If IsNumeric(Mid$(strlinea, pos1, pos2 - pos1)) Then
        NroEscala = Mid$(strlinea, pos1, pos2 - pos1)
    Else
        InsertaError 1, 3
        HuboError = True
        Exit Sub
    End If
    
    'Coord1
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Coord1 = Mid(strlinea, pos1, pos2 - pos1)
    If IsNumeric(Coord1) Then
       If CLng(Coord1) = 0 Then
          Coord1 = "null"
       End If
    Else
       InsertaError 2, 3
       HuboError = True
       Exit Sub
    End If
    
    'Coord2
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Coord2 = Mid(strlinea, pos1, pos2 - pos1)
    If IsNumeric(Coord2) Then
       If CLng(Coord2) = 0 Then
          Coord2 = "null"
       End If
    Else
       InsertaError 3, 3
       HuboError = True
       Exit Sub
    End If
    
    'Coord3
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Coord3 = Mid(strlinea, pos1, pos2 - pos1)
    If IsNumeric(Coord3) Then
       If CLng(Coord3) = 0 Then
          Coord3 = "null"
       End If
    Else
       InsertaError 4, 3
       HuboError = True
       Exit Sub
    End If
    
    'Coord4
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Coord4 = Mid(strlinea, pos1, pos2 - pos1)
    If IsNumeric(Coord4) Then
       If CLng(Coord4) = 0 Then
          Coord4 = "null"
       End If
    Else
       InsertaError 5, 3
       HuboError = True
       Exit Sub
    End If
    
    'Coord5
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Coord5 = Mid(strlinea, pos1, pos2 - pos1)
    If IsNumeric(Coord5) Then
       If CLng(Coord5) = 0 Then
          Coord5 = "null"
       End If
    Else
       InsertaError 6, 3
       HuboError = True
       Exit Sub
    End If
    
    'Orden
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Orden = Mid(strlinea, pos1, pos2 - pos1)
    If Not IsNumeric(Orden) Then
       InsertaError 7, 3
       HuboError = True
       Exit Sub
    End If
               
    'Valor
    pos1 = pos2 + 1
    pos2 = Len(strlinea)
    Valor = Mid(strlinea, pos1, pos2)

' ====================================================================
'   Validar los parametros Levantados

'Que exista la escala
StrSql = "SELECT * FROM cabgrilla where cgrnro = " & NroEscala
OpenRecordset StrSql, rs_Escala
If rs_Escala.EOF Then
    Flog.Writeline "No se encontro la escala nro. " & NroEscala
    InsertaError 1, 8
    HuboError = True
    Exit Sub
End If

'Que el valor sea valido
ValorCorrecto = True
If Not IsNumeric(Valor) Then
    Flog.Writeline "El valor no es numerico " & Monto
    InsertaError 2, 5
    HuboError = True
    Exit Sub
Else
    ValorCorrecto = True
End If

'=============================================================
'Inserto la escala

'Controlo si ya existe
StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroEscala & " AND vgrorden = " & Orden

If Coord1 = "null" Then
   StrSql = StrSql & " AND vgrcoor_1 IS NULL "
Else
   StrSql = StrSql & " AND vgrcoor_1 = " & Coord1
End If

If Coord2 = "null" Then
   StrSql = StrSql & " AND vgrcoor_2 IS NULL "
Else
   StrSql = StrSql & " AND vgrcoor_2 = " & Coord2
End If

If Coord3 = "null" Then
   StrSql = StrSql & " AND vgrcoor_3 IS NULL "
Else
   StrSql = StrSql & " AND vgrcoor_3 = " & Coord3
End If

If Coord4 = "null" Then
   StrSql = StrSql & " AND vgrcoor_4 IS NULL "
Else
   StrSql = StrSql & " AND vgrcoor_4 = " & Coord4
End If

If Coord5 = "null" Then
   StrSql = StrSql & " AND vgrcoor_5 IS NULL "
Else
   StrSql = StrSql & " AND vgrcoor_5 = " & Coord5
End If

OpenRecordset StrSql, rs_ValEscala

If rs_ValEscala.EOF Then

    StrSql = "INSERT INTO valgrilla ("
    StrSql = StrSql & "cgrnro, vgrcoor_1, vgrcoor_2, vgrcoor_3, vgrcoor_4, vgrcoor_5, vgrorden, vgrvalor "
    StrSql = StrSql & ") VALUES (" & NroEscala
    StrSql = StrSql & "," & Coord1
    StrSql = StrSql & "," & Coord2
    StrSql = StrSql & "," & Coord3
    StrSql = StrSql & "," & Coord4
    StrSql = StrSql & "," & Coord5
    StrSql = StrSql & "," & Orden
    StrSql = StrSql & "," & Valor
    StrSql = StrSql & " )"
    
Else
    
    StrSql = " UPDATE valgrilla "
    StrSql = StrSql & " SET vgrvalor = " & Valor
    StrSql = StrSql & " WHERE cgrnro = " & NroEscala & " AND vgrorden = " & Orden
    
    If Coord1 = "null" Then
       StrSql = StrSql & " AND vgrcoor_1 IS NULL "
    Else
       StrSql = StrSql & " AND vgrcoor_1 = " & Coord1
    End If
    
    If Coord2 = "null" Then
       StrSql = StrSql & " AND vgrcoor_2 IS NULL "
    Else
       StrSql = StrSql & " AND vgrcoor_2 = " & Coord2
    End If
    
    If Coord3 = "null" Then
       StrSql = StrSql & " AND vgrcoor_3 IS NULL "
    Else
       StrSql = StrSql & " AND vgrcoor_3 = " & Coord3
    End If
    
    If Coord4 = "null" Then
       StrSql = StrSql & " AND vgrcoor_4 IS NULL "
    Else
       StrSql = StrSql & " AND vgrcoor_4 = " & Coord4
    End If
    
    If Coord5 = "null" Then
       StrSql = StrSql & " AND vgrcoor_5 IS NULL "
    Else
       StrSql = StrSql & " AND vgrcoor_5 = " & Coord5
    End If

End If

objConn.Execute StrSql, , adExecuteNoRecords

Flog.Writeline " Escala insertada "

Fin:
'cierro y libero
If rs_Escala.State = adStateOpen Then rs_Escala.Close
If rs_ValEscala.State = adStateOpen Then rs_ValEscala.Close

Set rs_Escala = Nothing
Set rs_ValEscala = Nothing

Exit Sub

Manejador_De_Error:
    HuboError = True

    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strlinea
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin

End Sub


Public Sub LineaModelo_250(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta en Acumuladores Mensuales.
' Autor      : Fernando Favre
'              El formato es:
'                   Legajo; Acunro; Monto; catidad; año; mes
' Fecha      : 10/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1 As Long
Dim pos2 As Long
Dim i As Long

Dim Tercero As Long
Dim NroLegajo As Long
Dim acunro As Long
Dim Monto As Single
Dim Cantidad As Single
Dim Anio As Long
Dim mes As Long
Dim PliqNro As Long

Dim rs_Empleado As New ADODB.Recordset
Dim rs_Acumulador As New ADODB.Recordset
Dim rs_Acu_Mes As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset

On Error GoTo Manejador_De_Error

' El formato es:
' Legajo; Acunro; Monto; cantidad; año; mes

    'Nro de Legajo
    pos1 = 1
    pos2 = InStr(pos1, strlinea, Separador)
    If IsNumeric(Mid$(strlinea, pos1, pos2 - pos1)) Then
        NroLegajo = Mid$(strlinea, pos1, pos2 - pos1)
    Else
        InsertaError 1, 8
        HuboError = True
        Exit Sub
    End If
    
    'Acumulador
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    acunro = Mid(strlinea, pos1, pos2 - pos1)

    'Monto
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Monto = Mid(strlinea, pos1, pos2 - pos1)
               
    'Cantidad
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Cantidad = Mid(strlinea, pos1, pos2 - pos1)
        
    'Año
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strlinea, Separador)
    Anio = Mid(strlinea, pos1, pos2 - pos1)
        
    'Mes
    pos1 = pos2 + 1
    pos2 = Len(strlinea)
    mes = Mid(strlinea, pos1, pos2)

' ====================================================================
'   Validar los parametros Levantados

'Que exista el legajo
StrSql = "SELECT * FROM empleado where empleg = " & NroLegajo
OpenRecordset StrSql, rs_Empleado
If rs_Empleado.EOF Then
    Flog.Writeline "No se encontro el legajo " & NroLegajo
    InsertaError 1, 8
    HuboError = True
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
    HuboError = True
    Exit Sub
End If

'Que el monto sea valido
If Not IsNumeric(Monto) Then
    Flog.Writeline "El monto no es numerico " & Monto
    InsertaError 3, 5
    HuboError = True
    Exit Sub
End If

'Que la cantidad sea valida
If Not IsNumeric(Cantidad) Then
    Flog.Writeline "La cantidad no es numerico " & Cantidad
    InsertaError 4, 5
    HuboError = True
    Exit Sub
End If

'Que el año sea valido
If Not IsNumeric(Anio) Then
    Flog.Writeline "El año no es numerico " & Anio
    InsertaError 5, 5
    HuboError = True
    Exit Sub
End If

'Que el mes sea valido
If Not IsNumeric(mes) Then
    Flog.Writeline "El mes no es numerico " & Anio
    InsertaError 6, 5
    HuboError = True
    Exit Sub
End If

'Busco el pliqnro correspondiente a ese año y mes
StrSql = "SELECT * FROM PERIODO WHERE pliqmes =" & mes
StrSql = StrSql & " AND pliqanio =" & Anio
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.Writeline "No existe periodo correspondiente al año " & Anio & " y mes  " & mes
    InsertaError 6, 5
    HuboError = True
    Exit Sub
Else
    PliqNro = rs_Periodo!PliqNro
End If

'=============================================================
'Busco si existe el acu_mes
StrSql = "SELECT * FROM acu_mes " & _
         " WHERE acunro = " & acunro & _
         " AND ammes = " & mes & _
         " AND amanio = " & Anio & _
         " AND ternro = " & Tercero
OpenRecordset StrSql, rs_Acu_Mes

If rs_Acu_Mes.EOF Then
        StrSql = "INSERT INTO acu_mes ("
        StrSql = StrSql & "acunro, ammonto, ammontoreal, amcant, ternro, amanio, ammes"
        StrSql = StrSql & ") VALUES (" & acunro
        StrSql = StrSql & "," & Monto
        StrSql = StrSql & "," & Monto
        StrSql = StrSql & "," & Cantidad
        StrSql = StrSql & "," & Tercero
        StrSql = StrSql & "," & Anio
        StrSql = StrSql & "," & mes
        StrSql = StrSql & " )"
        objConn.Execute StrSql, , adExecuteNoRecords
    
        Flog.Writeline "Acumulador insertado "
Else
        'Piso los datos del acu_mes
        StrSql = "UPDATE acu_mes SET ammonto = " & Monto
        StrSql = StrSql & " , amcant = " & Cantidad
        StrSql = StrSql & " , ammontoreal = " & Monto
        StrSql = StrSql & " WHERE acunro = " & acunro
        StrSql = StrSql & " AND amanio = " & Anio
        StrSql = StrSql & " AND ammes = " & mes
        StrSql = StrSql & " AND ternro = " & Tercero
        objConn.Execute StrSql, , adExecuteNoRecords
            
        Flog.Writeline "Acumulador Actualizado "
End If

Fin:
'cierro y libero
If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Acu_Mes.State = adStateOpen Then rs_Acu_Mes.Close
If rs_Acumulador.State = adStateOpen Then rs_Acumulador.Close

Set rs_Empleado = Nothing
Set rs_Periodo = Nothing
Set rs_Acu_Mes = Nothing
Set rs_Acumulador = Nothing
Exit Sub

Manejador_De_Error:
    HuboError = True

    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strlinea
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin

End Sub


