Attribute VB_Name = "MdlFuncionesImport"
Option Explicit

Public Sub Inicializar_Globales()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que inicializa todas las variables globales que interesan al proceso
' Autor      : FGZ
' Fecha      : 22/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

    Continuar = True
    Medida_Clase = ""
    Medida_Motivo = ""
    Accion = ""
    SubAccion = ""
    
    

End Sub



Public Sub Insertar_His_Estructura(ByVal Tenro As Long, ByVal Estrnro As Long, ByVal Tercero As Long, ByVal FechaDesde As String, ByVal FechaHasta As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que inserta en his_estructura la estructura del tipo especificada
'               en el rango de fechas especificado. Si ya existe una estructura del mismo tipo abierta ==> se cierra el dia anterior al la
'               fecha_desde y se abre la nueva estructura.
' Autor      : FGZ
' Fecha      : 14/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_His_Estructura As New ADODB.Recordset
Dim Fecha_Desde As Date
Dim Fecha_Hasta As Date

'On Error GoTo MELocal

Fecha_Desde = CDate(FechaDesde)
If Not EsNulo(FechaHasta) Then
    Fecha_Hasta = CDate(FechaHasta)
End If

    StrSql = " SELECT * FROM his_estructura "
    StrSql = StrSql & " WHERE ternro = " & Tercero
    StrSql = StrSql & " AND tenro = " & Tenro
    StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(Fecha_Desde) & ") "
    StrSql = StrSql & " AND ((" & ConvFecha(Fecha_Desde) & " <= htethasta) or (htethasta is null))"
    StrSql = StrSql & " ORDER BY htetdesde "
    If rs_His_Estructura.State = adStateOpen Then rs_His_Estructura.Close
    OpenRecordset StrSql, rs_His_Estructura
    If Not rs_His_Estructura.EOF Then
        rs_His_Estructura.MoveLast
        If rs_His_Estructura!Estrnro = Estrnro Then
            'la estructura es la misma, reviso las fechas
            If Fecha_Desde < rs_His_Estructura!htetdesde Then
                StrSql = "UPDATE his_estructura SET htetdesde = " & ConvFecha(CDate(Fecha_Desde))
                StrSql = StrSql & " WHERE ternro = " & Tercero
                StrSql = StrSql & " AND tenro = " & Tenro
                StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(Fecha_Desde) & ") "
                StrSql = StrSql & " AND ((" & ConvFecha(Fecha_Desde) & " <= htethasta) or (htethasta is null))"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
            If Not EsNulo(FechaHasta) Then
                If Not EsNulo(rs_His_Estructura!htethasta) Then
                    If Fecha_Hasta < rs_His_Estructura!htethasta Then
                        StrSql = "UPDATE his_estructura SET htethasta = " & ConvFecha(CDate(Fecha_Hasta))
                        StrSql = StrSql & " WHERE ternro = " & Tercero
                        StrSql = StrSql & " AND tenro = " & Tenro
                        StrSql = StrSql & " AND estrnro = " & Estrnro
                        StrSql = StrSql & " AND htetdesde = " & ConvFecha(rs_His_Estructura!htetdesde)
                        StrSql = StrSql & " AND htethasta = " & ConvFecha(rs_His_Estructura!htethasta)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        'no actualizo porque la que estaba abarca mayor rango
                        FlogE.writeline Espacios(Tabulador * 3) & "Estructura no insertada. No actualizo porque la que estaba abarca mayor rango"
                    End If
                Else
                    StrSql = "UPDATE his_estructura SET htethasta = " & ConvFecha(CDate(Fecha_Hasta))
                    StrSql = StrSql & " WHERE ternro = " & Tercero
                    StrSql = StrSql & " AND estrnro = " & Estrnro
                    StrSql = StrSql & " AND tenro = " & Tenro
                    StrSql = StrSql & " AND htetdesde = " & ConvFecha(rs_His_Estructura!htetdesde)
                    StrSql = StrSql & " AND htethasta is null"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            Else
                If Not EsNulo(rs_His_Estructura!htethasta) Then
                    StrSql = "UPDATE his_estructura SET htethasta = NULL "
                    StrSql = StrSql & " WHERE ternro = " & Tercero
                    StrSql = StrSql & " AND tenro = " & Tenro
                    StrSql = StrSql & " AND estrnro = " & Estrnro
                    StrSql = StrSql & " AND htetdesde = " & ConvFecha(rs_His_Estructura!htetdesde)
                    StrSql = StrSql & " AND htethasta = " & ConvFecha(rs_His_Estructura!htethasta)
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
'                    StrSql = "UPDATE his_estructura SET htethasta = NULL "
'                    StrSql = StrSql & " WHERE ternro = " & Tercero
'                    StrSql = StrSql & " AND tenro = " & Tenro
'                    StrSql = StrSql & " AND estrnro = " & Estrnro
'                    StrSql = StrSql & " AND htetdesde = " & ConvFecha(rs_His_Estructura!htetdesde)
'                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
        Else    'tiene abierta otra estructura, la cierro y le abro la nueva
            If Fecha_Desde = rs_His_Estructura!htetdesde Then
                StrSql = "UPDATE his_estructura SET estrnro = " & Estrnro
                If Not EsNulo(FechaHasta) Then
                    StrSql = StrSql & " ,htethasta = " & ConvFecha(CDate(Fecha_Hasta))
                Else
                    StrSql = StrSql & " ,htethasta = null"
                End If
                StrSql = StrSql & " WHERE ternro = " & Tercero
                StrSql = StrSql & " AND tenro = " & Tenro
                StrSql = StrSql & " AND (htetdesde = " & ConvFecha(Fecha_Desde) & ") "
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                StrSql = "UPDATE his_estructura SET htethasta = " & ConvFecha(CDate(Fecha_Desde - 1))
                StrSql = StrSql & " WHERE ternro = " & Tercero
                StrSql = StrSql & " AND tenro = " & Tenro
                StrSql = StrSql & " AND estrnro = " & rs_His_Estructura!Estrnro
                StrSql = StrSql & " AND htetdesde = " & ConvFecha(rs_His_Estructura!htetdesde)
                If Not EsNulo(rs_His_Estructura!htethasta) Then
                    StrSql = StrSql & " AND htethasta = " & ConvFecha(rs_His_Estructura!htethasta)
                Else
                    StrSql = StrSql & " AND htethasta is null"
                End If
                objConn.Execute StrSql, , adExecuteNoRecords
            
                'Inserto la nueva estructura
                If Not EsNulo(FechaHasta) Then
                    StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde,htethasta) VALUES("
                    StrSql = StrSql & Tercero & "," & Estrnro & "," & Tenro & "," & ConvFecha(Fecha_Desde) & "," & ConvFecha(Fecha_Hasta) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde) VALUES("
                    StrSql = StrSql & Tercero & "," & Estrnro & "," & Tenro & "," & ConvFecha(Fecha_Desde) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
        End If
    Else
        'Inserto la nueva estructura
        If Not EsNulo(FechaHasta) Then
            StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde,htethasta) VALUES("
            StrSql = StrSql & Tercero & "," & Estrnro & "," & Tenro & "," & ConvFecha(Fecha_Desde) & "," & ConvFecha(Fecha_Hasta) & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde) VALUES("
            StrSql = StrSql & Tercero & "," & Estrnro & "," & Tenro & "," & ConvFecha(Fecha_Desde) & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    Call InsertarLogcambio(Tenro, Estrnro, Estrnro, FechaDesde, FechaHasta)
'Exit Sub
'MELocal:
'    Resume Next
End Sub



Public Function Baja_Empleado(ByVal ternro As Long, ByVal FechaBaja, ByVal Causa As Long) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que realiza la baja del empleado (ternro)
'              Cierra fase dependiendo del parametro pasado en FechaBaja
' Autor      : FGZ
' Fecha      : 03/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Estrnro As Long

Dim rs_Fases As New ADODB.Recordset
Dim rs_Causa As New ADODB.Recordset
Dim rs_Causa_Sitrev As New ADODB.Recordset

   
    StrSql = "SELECT * FROM fases WHERE empleado = " & ternro
    StrSql = StrSql & " AND altfec <= " & ConvFecha(FechaBaja)
    If rs_Fases.State = adStateOpen Then rs_Fases.Close
    OpenRecordset StrSql, rs_Fases
    If Not rs_Fases.EOF Then
        rs_Fases.MoveLast
        'If CBool(rs_Fases!estado) Then
            StrSql = "UPDATE fases SET estado = 0, bajfec=" & ConvFecha(FechaBaja)
            StrSql = StrSql & ", caunro=" & Causa
        'Else
        '    StrSql = "UPDATE fases SET caunro = " & Causa
        'End If
        StrSql = StrSql & " WHERE fasnro =" & rs_Fases!fasnro
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
        
    
    StrSql = "SELECT caudesvin FROM causa WHERE caunro = " & Causa
    If rs_Causa.State = adStateOpen Then rs_Causa.Close
    OpenRecordset StrSql, rs_Causa
    If Not rs_Causa.EOF Then
        'Si la causa de baja indicada tiene la marca de desvinculación en true:
        '   ==> se debe colocar empest = false
        '   sino no hacer nada
        If CBool(rs_Causa!caudesvin) Then
            StrSql = "UPDATE empleado SET empest=0 WHERE ternro = " & ternro
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        '=========================================================================================
        'VERIFICAR SI TIENE UNA SIT. DE REVISTA RELACIONADA AL CAUNRO
        'SI ES ASI, HAY QUE CVREAR UNA HIS_ESTRUCTURA CON ESE SIT.REVISTA
        StrSql = "SELECT estrnro, caunro "
        StrSql = StrSql & "FROM causa_sitrev "
        StrSql = StrSql & "WHERE caunro = " & Causa
        If rs_Causa_Sitrev.State = adStateOpen Then rs_Causa_Sitrev.Close
        OpenRecordset StrSql, rs_Causa_Sitrev
        If Not rs_Causa_Sitrev.EOF Then
            Estrnro = rs_Causa_Sitrev!Estrnro
        Else
            Estrnro = 0
        End If
    
        ' Esta relacionado a una situacion de revista
        If Estrnro <> 0 Then
            'Cierro cualquier estructura abierta --------------------------------------------
            StrSql = "UPDATE his_estructura SET"
            StrSql = StrSql & " htethasta = " & ConvFecha(DateAdd("d", -1, CDate(FechaBaja)))
            'StrSql = StrSql & " htethasta = " & ConvFecha(CDate(FechaBaja))
            StrSql = StrSql & " WHERE tenro = 30 "
            StrSql = StrSql & " AND ternro = " & ternro
            StrSql = StrSql & " AND htethasta IS NULL "
            objConn.Execute StrSql, , adExecuteNoRecords

            'Crear el his_estructura --------------------------------------------
            StrSql = "INSERT INTO his_estructura "
            StrSql = StrSql & " (tenro, ternro, estrnro, htetdesde,htethasta) "
            StrSql = StrSql & " VALUES (30, " & ternro & ", "
            StrSql = StrSql & Estrnro & ", "
            StrSql = StrSql & ConvFecha(FechaBaja) & ", NULL) "
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        '=========================================================================================

    End If
    Baja_Empleado = True

'cierro y libero todo
If rs_Fases.State = adStateOpen Then rs_Fases.Close
If rs_Causa.State = adStateOpen Then rs_Causa.Close
If rs_Causa_Sitrev.State = adStateOpen Then rs_Causa_Sitrev.Close
    
Set rs_Fases = Nothing
Set rs_Causa = Nothing
Set rs_Causa_Sitrev = Nothing
End Function




Public Sub Insertar_Documento(ByVal Doc As String, ByVal tipoDoc As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que inserta el documento del empleado
' Autor      : FGZ
' Fecha      : 07/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_TDoc     As New ADODB.Recordset
Dim rs          As New ADODB.Recordset

Dim Masculino   As Boolean
Dim Argentino   As Boolean
Dim Cuil        As String
Dim tipo        As String
Dim Genero_Cuil_Auto As Boolean

If tipoDoc <> 0 Then
    Doc = Format_Str(Replace(Doc, ".", ""), 30, False, "")
    
    StrSql = "SELECT * FROM ter_doc  "
    StrSql = StrSql & " WHERE ter_doc.tidnro = " & tipoDoc '& " AND ter_doc.nrodoc = '" & Doc & "'"
    StrSql = StrSql & " AND ternro = " & Empleado.Tercero
    OpenRecordset StrSql, rs_TDoc
        
    If rs_TDoc.EOF Then
        StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
        StrSql = StrSql & " VALUES(" & Empleado.Tercero & "," & tipoDoc & ",'" & Doc & "')"
    Else 'Actualizo
        StrSql = " UPDATE ter_doc SET "
        StrSql = StrSql & " nrodoc = '" & Doc & "'"
        StrSql = StrSql & " WHERE ternro = " & Empleado.Tercero
        StrSql = StrSql & " AND tidnro = " & tipoDoc
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
    Genero_Cuil_Auto = False
    If Genero_Cuil_Auto Then
        'verifico si ya tiene CUIL
        StrSql = "SELECT * FROM ter_doc  "
        StrSql = StrSql & " WHERE ter_doc.tidnro = 10"
        StrSql = StrSql & " AND ternro = " & Empleado.Tercero
        OpenRecordset StrSql, rs_TDoc
            
        If rs_TDoc.EOF Then
            'reviso si puedo generarlo automaticamente
            'unicamente lo genero si es Argentino y el tipo de doc es LC, LE o DNI
            If tipoDoc = 1 Or tipoDoc = 2 Or tipoDoc = 3 Then
                Select Case tipoDoc
                Case 1:
                    tipo = "DNI"
                Case 2:
                    tipo = "LE"
                Case 3:
                    tipo = "LC"
                Case Else
                    tipo = ""
                End Select
                
                'Busco el sexo del legajo
                StrSql = " SELECT * FROM tercero WHERE ternro = " & Empleado.Tercero
                If rs.State = adStateOpen Then rs.Close
                OpenRecordset StrSql, rs
                If rs.EOF Then
                    Flog.writeline Espacios(Tabulador * 3) & "No se encuentra el legajo." & Empleado.Legajo
                Else
                    Masculino = CBool(rs!Tersex)
                End If
                
                'busco la nacionalidad del legajo
                StrSql = "SELECT * FROM tercero "
                StrSql = StrSql & " INNER JOIN nacionalidad ON tercero.nacionalnro = nacionalidad.nacionalnro AND nacionalidad.nacionaldefault = -1"
                StrSql = StrSql & " WHERE ternro = " & Empleado.Tercero
                If rs.State = adStateOpen Then rs.Close
                OpenRecordset StrSql, rs
                If rs.EOF Then
                    Argentino = False
                Else
                    Argentino = True
                End If
                
                If Argentino Then
                    Cuil = New_Generar_Cuil(tipo, Doc, Masculino)
                    
                    StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
                    StrSql = StrSql & " VALUES(" & Empleado.Tercero & ",10,'" & Cuil & "')"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    Flog.writeline Espacios(Tabulador * 3) & "CUIL generado Automaticamente."
                Else
                    Flog.writeline Espacios(Tabulador * 3) & "El empleado no es Argentino, no se generará el CUIL."
                End If
            Else
                'unicamente lo genero si es Argentino y el tipo de doc es LC, LE o DNI
            End If
        Else 'Actualizo
            Flog.writeline Espacios(Tabulador * 3) & "ya tiene CUIL."
        End If
    End If
End If

'cierro y libero
If rs_TDoc.State = adStateOpen Then rs_TDoc.Close
Set rs_TDoc = Nothing
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing
End Sub

Public Sub Insertar_Documento_Familiar(ByVal Tercero As Long, ByVal Doc As String, ByVal tipoDoc As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que inserta el documento del empleado
' Autor      : FGZ
' Fecha      : 07/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_TDoc     As New ADODB.Recordset
Dim rs          As New ADODB.Recordset

Dim Masculino   As Boolean
Dim Argentino   As Boolean
Dim Cuil        As String
Dim tipo        As String

If tipoDoc <> 0 Then
    Doc = Format_Str(Replace(Doc, ".", ""), 30, False, "")
    
    StrSql = "SELECT * FROM ter_doc  "
    StrSql = StrSql & " WHERE ter_doc.tidnro = " & tipoDoc '& " AND ter_doc.nrodoc = '" & Doc & "'"
    StrSql = StrSql & " AND ternro = " & Tercero
    OpenRecordset StrSql, rs_TDoc
        
    If rs_TDoc.EOF Then
        StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
        StrSql = StrSql & " VALUES(" & Tercero & "," & tipoDoc & ",'" & Doc & "')"
    Else 'Actualizo
        StrSql = " UPDATE ter_doc SET "
        StrSql = StrSql & " nrodoc = '" & Doc & "'"
        StrSql = StrSql & " WHERE ternro = " & Tercero
        StrSql = StrSql & " AND tidnro = " & tipoDoc
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'verifico si ya tiene CUIL
    StrSql = "SELECT * FROM ter_doc  "
    StrSql = StrSql & " WHERE ter_doc.tidnro = 10"
    StrSql = StrSql & " AND ternro = " & Tercero
    OpenRecordset StrSql, rs_TDoc
        
    If rs_TDoc.EOF Then
        'reviso si puedo generarlo automaticamente
        'unicamente lo genero si es Argentino y el tipo de doc es LC, LE o DNI
        If tipoDoc = 1 Or tipoDoc = 2 Or tipoDoc = 3 Then
            Select Case tipoDoc
            Case 1:
                tipo = "DNI"
            Case 2:
                tipo = "LE"
            Case 3:
                tipo = "LC"
            Case Else
                tipo = ""
            End Select
            
            'Busco el sexo del legajo
            StrSql = " SELECT * FROM tercero WHERE ternro = " & Tercero
            If rs.State = adStateOpen Then rs.Close
            OpenRecordset StrSql, rs
            If rs.EOF Then
                Flog.writeline Espacios(Tabulador * 3) & "No se encuentra el legajo." & Empleado.Legajo
            Else
                Masculino = CBool(rs!Tersex)
            End If
            
            'busco la nacionalidad del legajo
            StrSql = "SELECT * FROM tercero "
            StrSql = StrSql & " INNER JOIN nacionalidad ON tercero.nacionalnro = nacionalidad.nacionalnro AND nacionalidad.nacionaldefault = -1"
            StrSql = StrSql & " WHERE ternro = " & Tercero
            If rs.State = adStateOpen Then rs.Close
            OpenRecordset StrSql, rs
            If rs.EOF Then
                Argentino = False
            Else
                Argentino = True
            End If
            
            If Argentino Then
                Cuil = New_Generar_Cuil(tipo, Doc, Masculino)
                
                StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
                StrSql = StrSql & " VALUES(" & Tercero & ",10,'" & Cuil & "')"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Flog.writeline Espacios(Tabulador * 3) & "CUIL generado Automaticamente."
            Else
                Flog.writeline Espacios(Tabulador * 3) & "El empleado no es Argentino, no se generará el CUIL."
            End If
        Else
            'unicamente lo genero si es Argentino y el tipo de doc es LC, LE o DNI
        End If
    Else 'Actualizo
        Flog.writeline Espacios(Tabulador * 3) & "ya tiene CUIL."
    End If
End If

'cierro y libero
If rs_TDoc.State = adStateOpen Then rs_TDoc.Close
Set rs_TDoc = Nothing
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing
End Sub


Public Sub Insertar_Mail(ByVal Mail As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que actualiza el e-mail del empleado
' Autor      : FGZ
' Fecha      : 07/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
        StrSql = " UPDATE empleado SET "
        StrSql = StrSql & " empemail = '" & Format_Str(Mail, 100, False, "") & "'"
        StrSql = StrSql & " WHERE ternro = " & Empleado.Tercero
        objConn.Execute StrSql, , adExecuteNoRecords
End Sub


Public Sub Insertar_Telefono(ByVal Telefono As String, ByVal Celular As Boolean, ByVal Default As Boolean, ByVal Fax As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que inserta el telefono del empleado
' Autor      : FGZ
' Fecha      : 07/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim TipoDomi As Long
Dim NroDom As Long

Dim rs_Tel As New ADODB.Recordset
Dim rs_CabDom As New ADODB.Recordset


    Telefono = Format_Str(Telefono, 20, False, "")
    TipoDomi = 2

    StrSql = " SELECT * FROM cabdom "
    StrSql = StrSql & " WHERE tipnro = " & TipoDomi & " AND domdefault = -1 AND tidonro = 2 "
    StrSql = StrSql & " AND ternro = " & Empleado.Tercero
    If rs_CabDom.State = adStateOpen Then rs_CabDom.Close
    OpenRecordset StrSql, rs_CabDom
    If rs_CabDom.EOF Then
        StrSql = " INSERT INTO cabdom(tipnro,ternro,domdefault,tidonro) "
        StrSql = StrSql & " VALUES(" & TipoDomi & "," & Empleado.Tercero & ",-1,2)"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        NroDom = getLastIdentity(objConn, "cabdom")
        
        If Telefono <> "" And Default Then
          StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular) "
          StrSql = StrSql & " VALUES(" & NroDom & ",'" & Telefono & "',0,-1,0)"
          objConn.Execute StrSql, , adExecuteNoRecords
        Else
            If Telefono <> "" And Celular Then
                  StrSql = "SELECT * FROM telefono "
                  StrSql = StrSql & " WHERE domnro =" & NroDom
                  StrSql = StrSql & " AND telnro ='" & Telefono & "'"
                  If rs_Tel.State = adStateOpen Then rs_Tel.Close
                  OpenRecordset StrSql, rs_Tel
                  If rs_Tel.EOF Then
                      StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular) "
                      StrSql = StrSql & " VALUES(" & NroDom & ",'" & Telefono & "',0,0,-1)"
                      objConn.Execute StrSql, , adExecuteNoRecords
                  End If
            Else
                If Telefono <> "" Then
                  StrSql = "SELECT * FROM telefono "
                  StrSql = StrSql & " WHERE domnro =" & NroDom
                  StrSql = StrSql & " AND telnro ='" & Telefono & "'"
                  If rs_Tel.State = adStateOpen Then rs_Tel.Close
                  OpenRecordset StrSql, rs_Tel
                  If rs_Tel.EOF Then
                      StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular) "
                      StrSql = StrSql & " VALUES(" & NroDom & ",'" & Telefono & "',0,0,0)"
                      objConn.Execute StrSql, , adExecuteNoRecords
                  End If
                End If
            End If
        End If
    Else
        NroDom = rs_CabDom!Domnro
      
        If Telefono <> "" And Default Then
            StrSql = " UPDATE telefono SET "
            StrSql = StrSql & " telnro = '" & Telefono & "'"
            StrSql = StrSql & " WHERE domnro = " & NroDom
            StrSql = StrSql & " AND teldefault = -1 "
            StrSql = StrSql & " AND telcelular = 0 "
            StrSql = StrSql & " AND telfax = 0 "
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            If Telefono <> "" And Celular Then
                StrSql = " UPDATE telefono SET "
                StrSql = StrSql & " telnro = '" & Telefono & "'"
                StrSql = StrSql & " WHERE domnro = " & NroDom
                StrSql = StrSql & " AND teldefault = 0 "
                StrSql = StrSql & " AND telcelular = -1 "
                StrSql = StrSql & " AND telfax = 0 "
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                If Telefono <> "" Then
                    StrSql = " UPDATE telefono SET "
                    StrSql = StrSql & " telnro = '" & Telefono & "'"
                    StrSql = StrSql & " WHERE domnro = " & NroDom
                    StrSql = StrSql & " AND teldefault = 0 "
                    StrSql = StrSql & " AND telcelular = 0 "
                    StrSql = StrSql & " AND telfax = 0 "
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
        End If
    End If

'cierro y libero
If rs_Tel.State = adStateOpen Then rs_Tel.Close
If rs_CabDom.State = adStateOpen Then rs_CabDom.Close
Set rs_Tel = Nothing
Set rs_CabDom = Nothing
End Sub


Public Sub Insertar_Fecha(ByVal Clase As String, ByVal Fecha As Date, ByVal Fecha_Desde, ByVal Fecha_Hasta)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que inserta la fecha correspondiente al tipo
' Autor      : FGZ
' Fecha      : 07/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Fases As New ADODB.Recordset
Dim FechaDesde As Date
Dim FechaHasta As Date


FechaDesde = CDate(Fecha_Desde)
If Not EsNulo(Fecha_Hasta) Then
    FechaHasta = CDate(Fecha_Hasta)
End If

Select Case UCase(Clase)
Case "01":  'Fecha de Ingreso del empleado
        StrSql = " UPDATE empleado SET "
        StrSql = StrSql & " empfecalta = " & ConvFecha(Fecha)
        'StrSql = StrSql & ",empfecbaja = " & ConvFecha(Fecha)
        'StrSql = StrSql & ",empfaltagr = " & ConvFecha(fecha)
        StrSql = StrSql & " WHERE ternro = " & Empleado.Tercero
        objConn.Execute StrSql, , adExecuteNoRecords
    
        StrSql = "SELECT * FROM fases "
        StrSql = StrSql & " where empleado =" & Empleado.Tercero
        StrSql = StrSql & " and altfec = " & ConvFecha(Fecha)
        OpenRecordset StrSql, rs_Fases
        If rs_Fases.EOF Then
            'Inserto la Fase
            StrSql = " INSERT INTO fases("
            StrSql = StrSql & "empleado,altfec,estado,sueldo,vacaciones,indemnizacion,real,fasrecofec"
            If Not EsNulo(Fecha_Hasta) Then
                StrSql = StrSql & ",bajfec"
            End If
            StrSql = StrSql & ") VALUES( "
            StrSql = StrSql & Empleado.Tercero
            StrSql = StrSql & "," & ConvFecha(Fecha)
            StrSql = StrSql & ",-1,-1,-1,-1,-1,-1"
            If Not EsNulo(Fecha_Hasta) Then
                StrSql = StrSql & "," & ConvFecha(FechaHasta)
            End If
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            If Not EsNulo(Fecha_Hasta) Then
                StrSql = " UPDATE fases SET "
                StrSql = StrSql & " bajfec = " & ConvFecha(Fecha)
                StrSql = StrSql & " WHERE fasnro =" & rs_Fases!fasnro
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
Case "Z3":  'Fecha de baja del empleado
        
        StrSql = " UPDATE empleado SET "
        'StrSql = StrSql & " empfecalta = " & ConvFecha(Fecha)
        StrSql = StrSql & " empfecbaja = " & ConvFecha(Fecha)
        'StrSql = StrSql & ",empfaltagr = " & ConvFecha(fecha)
        StrSql = StrSql & " WHERE ternro = " & Empleado.Tercero
        objConn.Execute StrSql, , adExecuteNoRecords
    
        'Inserto la Fase
        StrSql = "SELECT * FROM fases "
        'StrSql = StrSql & " WHERE estado = -1"
        StrSql = StrSql & " WHERE empleado = " & Empleado.Tercero
        StrSql = StrSql & " ORDER BY altfec "
        OpenRecordset StrSql, rs_Fases
        
        If Not rs_Fases.EOF Then
            rs_Fases.MoveLast
            
            StrSql = "UPDATE fases SET "
            StrSql = StrSql & " estado = 0"
            StrSql = StrSql & " ,bajfec = " & ConvFecha(Fecha)
            StrSql = StrSql & " WHERE fasnro = " & rs_Fases!fasnro
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            'No se deberia dar
        End If
Case Else
    'Flog.Writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
    Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ": No se va a tener en cuenta el tipo de fecha " & Clase
End Select

'cierro y libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing
End Sub


Public Sub Insertar_Novedad(ByVal Nomina As String, ByVal Monto As Double, Cantidad As Double, ByVal Fecha_Desde As String, ByVal Fecha_Hasta As String, ByVal Texto As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que mapea un cc-nomina en un par concnro-tparnro y
'               lo inserta en novemp. si no encuentra el mapeo ==> no se inserta nada.
' Autor      : FGZ
' Fecha      : 07/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim i As Integer
Dim concnro As Long
Dim tpanro As Long
Dim EsMonto As Boolean  'True = Monto False = cantidad
Dim Valor As Single
Dim FechaDesde As Date
Dim FechaHasta As Date

Dim rs_NovEmp As New ADODB.Recordset

FechaDesde = CDate(Fecha_Desde)
If Not EsNulo(Fecha_Hasta) Then
    FechaHasta = CDate(Fecha_Hasta)
End If

'Inserta la novedad para el monto
EsMonto = True
For i = 1 To 2
    Call CalcularMapeoNomina(Nomina, EsMonto, concnro, tpanro)
    'Call CalcularMapeoNominaAutomatico(Nomina, FechaDesde, concnro, tpanro)
    If concnro <> 0 And tpanro <> 0 Then
        If EsMonto Then
            Valor = Monto
        Else
            Valor = Cantidad
        End If
        
        StrSql = "SELECT * FROM novemp WHERE "
        StrSql = StrSql & " concnro = " & concnro
        StrSql = StrSql & " AND tpanro = " & tpanro
        StrSql = StrSql & " AND empleado = " & Empleado.Tercero
        StrSql = StrSql & " AND (nevigencia = 0 "
        StrSql = StrSql & " OR (nevigencia = -1 "
        If Not EsNulo(Fecha_Hasta) Then
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
            'A lo sumo va a actualizar una sola
            Do While Not rs_NovEmp.EOF
                If Not CBool(rs_NovEmp!nevigencia) Then
                    Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
                    Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ": No se puede insertar la novedad porque ya existe una sin vigencia"
                Else
                    If rs_NovEmp!nedesde = FechaDesde Then
                        If EsNulo(rs_NovEmp!neHasta) Then
'                            'es una delimitacion de la que ya existe
'                            StrSql = "UPDATE novemp SET "
'                            If EsNulo(Fecha_Hasta) Then
'                               StrSql = StrSql & " nehasta = null "
'                            Else
'                               StrSql = StrSql & " nehasta = " & ConvFecha(FechaHasta)
'                            End If
'                            StrSql = StrSql & " ,nevalor = " & Valor
'                            StrSql = StrSql & " WHERE nenro = " & rs_NovEmp!nenro
''                            StrSql = StrSql & " WHERE concnro = " & concnro
''                            StrSql = StrSql & " AND tpanro = " & tpanro
''                            StrSql = StrSql & " AND empleado = " & Empleado.Tercero
''                            StrSql = StrSql & " AND nedesde = " & ConvFecha(rs_NovEmp!nedesde)
'                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            'ya la tengo ==> actualizo el monto y la fecha hasta
                            If Not EsNulo(rs_NovEmp!neHasta) Then
                                If Not EsNulo(Fecha_Hasta) Then
                                    If Fecha_Hasta < rs_NovEmp!neHasta Then
                                        StrSql = "UPDATE novemp SET "
                                        StrSql = StrSql & " nehasta = " & ConvFecha(FechaHasta)
                                        StrSql = StrSql & " ,nevalor = " & Valor
                                        StrSql = StrSql & " WHERE nenro = " & rs_NovEmp!nenro
                                        objConn.Execute StrSql, , adExecuteNoRecords
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If rs_NovEmp!nedesde < FechaDesde Then
                            'es la nueva vigencia, actualizo la anterior e Inserto la nueva
                            StrSql = "UPDATE novemp SET "
                            'If EsNulo(Fecha_Hasta) Then
                               StrSql = StrSql & " nehasta = " & ConvFecha(FechaDesde - 1)
                            'Else
                            '   StrSql = StrSql & " nehasta = " & ConvFecha(FechaHasta)
                            'End If
                            'StrSql = StrSql & " ,nevalor = " & Valor
                            StrSql = StrSql & " WHERE nenro = " & rs_NovEmp!nenro
                            objConn.Execute StrSql, , adExecuteNoRecords
                        
                            'es la nueva vigencia, Inserto
                            StrSql = "INSERT INTO novemp ("
                            StrSql = StrSql & "empleado,concnro,tpanro,nevalor,nevigencia,nedesde"
                            If Not EsNulo(Fecha_Hasta) Then
                                StrSql = StrSql & ",nehasta"
                            End If
                            StrSql = StrSql & " ,netexto"
                            StrSql = StrSql & ") VALUES (" & Empleado.Tercero
                            StrSql = StrSql & "," & concnro
                            StrSql = StrSql & "," & tpanro
                            StrSql = StrSql & "," & Valor
                            StrSql = StrSql & ",-1"
                            StrSql = StrSql & "," & ConvFecha(FechaDesde)
                            If Not EsNulo(Fecha_Hasta) Then
                                StrSql = StrSql & "," & ConvFecha(FechaHasta)
                            End If
                            StrSql = StrSql & ",'" & Texto & "'"
                            StrSql = StrSql & " )"
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                End If
                rs_NovEmp.MoveNext
            Loop
        Else
                StrSql = "INSERT INTO novemp ("
                StrSql = StrSql & "empleado,concnro,tpanro,nevalor,nevigencia,nedesde"
                If Not EsNulo(Fecha_Hasta) Then
                    StrSql = StrSql & ",nehasta"
                End If
                StrSql = StrSql & " ,netexto"
                
                StrSql = StrSql & ") VALUES (" & Empleado.Tercero
                StrSql = StrSql & "," & concnro
                StrSql = StrSql & "," & tpanro
                StrSql = StrSql & "," & Valor
                StrSql = StrSql & ",-1"
                StrSql = StrSql & "," & ConvFecha(FechaDesde)
                If Not EsNulo(Fecha_Hasta) Then
                    StrSql = StrSql & "," & ConvFecha(FechaHasta)
                End If
                StrSql = StrSql & ",'" & Texto & "'"
                StrSql = StrSql & " )"
                objConn.Execute StrSql, , adExecuteNoRecords
        End If
        Call InsertarLogNovedad(Nomina, concnro, tpanro, EsMonto, Valor, Fecha_Desde, Fecha_Hasta, Texto)
    End If
    
    EsMonto = False
    'Inserta la novedad para la cantidad
Next i

'cierro y libero
If rs_NovEmp.State = adStateOpen Then rs_NovEmp.Close
Set rs_NovEmp = Nothing
End Sub

Public Sub Insertar_Novedad_2(ByVal concnro As Long, ByVal tpanro As Long, ByVal Monto As Single, ByVal Fecha_Desde As String, ByVal Fecha_Hasta As String, ByVal Texto As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que inserta una novedad para el par concnro-tparnro.
'
' Autor      : FGZ
' Fecha      : 12/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Dim Concnro As Long
'Dim tpanro As Long
Dim EsMonto As Boolean  'True = Monto False = cantidad
Dim Valor As Single
Dim FechaDesde As Date
Dim FechaHasta As Date

Dim rs_NovEmp As New ADODB.Recordset

    FechaDesde = CDate(Fecha_Desde)
    If Not EsNulo(Fecha_Hasta) Then
        FechaHasta = CDate(Fecha_Hasta)
    End If

    EsMonto = True
    'Inserta la novedad para el monto
    If concnro <> 0 And tpanro <> 0 Then
        Valor = Monto
        
        StrSql = "SELECT * FROM novemp WHERE "
        StrSql = StrSql & " concnro = " & concnro
        StrSql = StrSql & " AND tpanro = " & tpanro
        StrSql = StrSql & " AND empleado = " & Empleado.Tercero
        StrSql = StrSql & " AND (nevigencia = 0 "
        StrSql = StrSql & " OR (nevigencia = -1 "
        If Not EsNulo(Fecha_Hasta) Then
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
            'A lo sumo va a actualizar una sola
            Do While Not rs_NovEmp.EOF
                If Not CBool(rs_NovEmp!nevigencia) Then
                    Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
                    Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ": No se puede insertar la novedad porque ya existe una sin vigencia"
                Else
                    If rs_NovEmp!nedesde = FechaDesde Then
                        If EsNulo(rs_NovEmp!neHasta) Then
                        
                        Else
                            'ya la tengo ==> actualizo el monto y la fecha hasta
                            If Not EsNulo(rs_NovEmp!neHasta) Then
                                If Not EsNulo(Fecha_Hasta) Then
                                    If Fecha_Hasta < rs_NovEmp!neHasta Then
                                        StrSql = "UPDATE novemp SET "
                                        StrSql = StrSql & " nehasta = " & ConvFecha(FechaHasta)
                                        StrSql = StrSql & " ,nevalor = " & Valor
                                        StrSql = StrSql & " WHERE nenro = " & rs_NovEmp!nenro
                                        objConn.Execute StrSql, , adExecuteNoRecords
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If rs_NovEmp!nedesde < FechaDesde Then
                            'es la nueva vigencia, actualizo la anterior e Inserto la nueva
                            StrSql = "UPDATE novemp SET "
                            StrSql = StrSql & " nehasta = " & ConvFecha(FechaDesde - 1)
                            StrSql = StrSql & " WHERE nenro = " & rs_NovEmp!nenro
                            objConn.Execute StrSql, , adExecuteNoRecords
                        
                            'es la nueva vigencia, Inserto
                            StrSql = "INSERT INTO novemp ("
                            StrSql = StrSql & "empleado,concnro,tpanro,nevalor,nevigencia,nedesde"
                            If Not EsNulo(Fecha_Hasta) Then
                                StrSql = StrSql & ",nehasta"
                            End If
                            StrSql = StrSql & " ,netexto"
                            StrSql = StrSql & ") VALUES (" & Empleado.Tercero
                            StrSql = StrSql & "," & concnro
                            StrSql = StrSql & "," & tpanro
                            StrSql = StrSql & "," & Valor
                            StrSql = StrSql & ",-1"
                            StrSql = StrSql & "," & ConvFecha(FechaDesde)
                            If Not EsNulo(Fecha_Hasta) Then
                                StrSql = StrSql & "," & ConvFecha(FechaHasta)
                            End If
                            StrSql = StrSql & ",'" & Texto & "'"
                            StrSql = StrSql & " )"
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                End If
                rs_NovEmp.MoveNext
            Loop
        Else
                StrSql = "INSERT INTO novemp ("
                StrSql = StrSql & "empleado,concnro,tpanro,nevalor,nevigencia,nedesde"
                If Not EsNulo(Fecha_Hasta) Then
                    StrSql = StrSql & ",nehasta"
                End If
                StrSql = StrSql & " ,netexto"
                
                StrSql = StrSql & ") VALUES (" & Empleado.Tercero
                StrSql = StrSql & "," & concnro
                StrSql = StrSql & "," & tpanro
                StrSql = StrSql & "," & Valor
                StrSql = StrSql & ",-1"
                StrSql = StrSql & "," & ConvFecha(FechaDesde)
                If Not EsNulo(Fecha_Hasta) Then
                    StrSql = StrSql & "," & ConvFecha(FechaHasta)
                End If
                StrSql = StrSql & ",'" & Texto & "'"
                StrSql = StrSql & " )"
                objConn.Execute StrSql, , adExecuteNoRecords
        End If
        Call InsertarLogNovedad("", concnro, tpanro, EsMonto, Valor, Fecha_Desde, Fecha_Hasta, Texto)
    End If
    
'cierro y libero
If rs_NovEmp.State = adStateOpen Then rs_NovEmp.Close
Set rs_NovEmp = Nothing
End Sub


Public Sub Insertar_Novedad_Ajuste(ByVal concepto As String, ByVal Monto As Single, ByVal Fecha_Desde As String, ByVal Fecha_Hasta As String, ByVal Texto As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que Inserta una novedad por ajuste para el concepto.
' Autor      : FGZ
' Fecha      : 12/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim concnro As Long
Dim FechaDesde As Date
Dim FechaHasta As Date

Dim rs_NovAju As New ADODB.Recordset

    FechaDesde = CDate(Fecha_Desde)
    If Not EsNulo(Fecha_Hasta) Then
        FechaHasta = CDate(Fecha_Hasta)
    End If

    StrSql = "SELECT * FROM concepto "
    StrSql = StrSql & " WHERE conccod = '" & concepto & "'"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        concnro = rs!concnro
    Else
        concnro = 0
    End If
    
    If concnro <> 0 Then
        StrSql = "SELECT * FROM novaju WHERE "
        StrSql = StrSql & " concnro = " & concnro
        StrSql = StrSql & " AND empleado = " & Empleado.Tercero
        StrSql = StrSql & " AND (navigencia = 0 "
        StrSql = StrSql & " OR (navigencia = -1 "
        If Not EsNulo(Fecha_Hasta) Then
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
            'A lo sumo va a actualizar una sola
            Do While Not rs_NovAju.EOF
                If Not CBool(rs_NovAju!navigencia) Then
                    Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
                    Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ": No se puede insertar la novedad porque ya existe una sin vigencia"
                Else
                    If rs_NovAju!nadesde = FechaDesde Then
                        If EsNulo(rs_NovAju!naHasta) Then
                        
                        Else
                            'ya la tengo ==> actualizo el monto y la fecha hasta
                            If Not EsNulo(rs_NovAju!naHasta) Then
                                If Not EsNulo(Fecha_Hasta) Then
                                    If Fecha_Hasta < rs_NovAju!naHasta Then
                                        StrSql = "UPDATE novaju SET "
                                        StrSql = StrSql & " nahasta = " & ConvFecha(FechaHasta)
                                        StrSql = StrSql & " ,navalor = " & Valor
                                        StrSql = StrSql & " WHERE nanro = " & rs_NovAju!nanro
                                        objConn.Execute StrSql, , adExecuteNoRecords
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If rs_NovAju!nadesde < FechaDesde Then
                            'es la nueva vigencia, actualizo la anterior e Inserto la nueva
                            StrSql = "UPDATE novaju SET "
                            StrSql = StrSql & " nehasta = " & ConvFecha(FechaDesde - 1)
                            StrSql = StrSql & " WHERE nanro = " & rs_NovAju!nanro
                            objConn.Execute StrSql, , adExecuteNoRecords
                        
                            'es la nueva vigencia, Inserto
                            StrSql = "INSERT INTO novaju ("
                            StrSql = StrSql & "empleado,concnro,nevalor,nevigencia,nedesde"
                            If Not EsNulo(Fecha_Hasta) Then
                                StrSql = StrSql & ",nehasta"
                            End If
                            StrSql = StrSql & " ,netexto"
                            StrSql = StrSql & ") VALUES (" & Empleado.Tercero
                            StrSql = StrSql & "," & concnro
                            StrSql = StrSql & "," & Monto
                            StrSql = StrSql & ",-1"
                            StrSql = StrSql & "," & ConvFecha(FechaDesde)
                            If Not EsNulo(Fecha_Hasta) Then
                                StrSql = StrSql & "," & ConvFecha(FechaHasta)
                            End If
                            StrSql = StrSql & ",'" & Texto & "'"
                            StrSql = StrSql & " )"
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                End If
                rs_NovAju.MoveNext
            Loop
        Else
            StrSql = "INSERT INTO novaju ("
            StrSql = StrSql & "empleado,concnro,navalor,navigencia,nadesde"
            If Not EsNulo(Fecha_Hasta) Then
                StrSql = StrSql & ",nahasta"
            End If
            StrSql = StrSql & " ,natexto"
            
            StrSql = StrSql & ") VALUES (" & Empleado.Tercero
            StrSql = StrSql & "," & concnro
            StrSql = StrSql & "," & Monto
            StrSql = StrSql & ",-1"
            StrSql = StrSql & "," & ConvFecha(FechaDesde)
            If Not EsNulo(Fecha_Hasta) Then
                StrSql = StrSql & "," & ConvFecha(FechaHasta)
            End If
            StrSql = StrSql & ",'" & Texto & "'"
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        Call InsertarLogNovedad("", concnro, 0, True, Monto, Fecha_Desde, Fecha_Hasta, Texto)
    End If
    
'cierro y libero
If rs_NovAju.State = adStateOpen Then rs_NovAju.Close
Set rs_NovAju = Nothing
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing
End Sub



Public Sub Insertar_Licencia(ByVal tipo As Long, ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByVal cFeriados As Integer, ByVal CHabiles As Integer, ByVal ternro As Long, ByVal Autorizadas As Boolean, ByRef OK As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que inserta la licencia y su correspondiente complemento.
' Autor      : FGZ
' Fecha      : 23/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim NroVac As Long
Dim rs_existeLic As New ADODB.Recordset
Dim rs_Periodos_Vac As New ADODB.Recordset

    On Error GoTo MELocal
    OK = True
    'reviso si ya existe esa Licencia
    StrSql = " SELECT * FROM emp_lic WHERE (empleado = " & ternro & _
             " ) AND (tdnro = " & tipo & ") " & _
             " AND eltipo = 1" & _
             " AND elfechadesde = " & ConvFecha(FechaDesde) & _
             " AND elfechahasta = " & ConvFecha(FechaHasta)
    OpenRecordset StrSql, rs_existeLic

    If rs_existeLic.EOF Then
        StrSql = "INSERT INTO emp_lic (elcantdias,elcantdiasfer,elcantdiashab,eldiacompleto,eltipo,elfechadesde,elfechahasta,elhoradesde,elhorahasta,tdnro,licestnro,empleado) VALUES ("
        StrSql = StrSql & CHabiles & ","
        
        StrSql = StrSql & cFeriados & ","
        StrSql = StrSql & CHabiles & ","
        StrSql = StrSql & "-1,"
        StrSql = StrSql & "1,"
        StrSql = StrSql & ConvFecha(FechaDesde) & ","
        StrSql = StrSql & ConvFecha(FechaHasta) & ","
        StrSql = StrSql & "null,"
        StrSql = StrSql & "null,"
        If Autorizadas Then
            StrSql = StrSql & "2,2,"
        Else
            StrSql = StrSql & "2,1,"
        End If
        StrSql = StrSql & ternro & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
            
            
        'De acurdo al tipo de Licencia puede que se necesite hacer otra cosa
        Select Case tipo
        Case 1:    'Licencia Gremial
        Case 2:    'Licencia por Vacaciones
            StrSql = "SELECT * FROM vacacion "
            StrSql = StrSql & " WHERE vacfecdesde <= " & ConvFecha(FechaHasta)
            StrSql = StrSql & " AND  vacfechasta >= " & ConvFecha(FechaDesde)
            StrSql = StrSql & " ORDER BY vacnro"
            OpenRecordset StrSql, rs_Periodos_Vac
        
            If Not rs_Periodos_Vac.EOF Then
                StrSql = "INSERT INTO lic_vacacion (emp_licnro,licvacmanual,vacnro) VALUES ("
                StrSql = StrSql & getLastIdentity(objConn, "emp_lic") & ","
                StrSql = StrSql & "0,"
                StrSql = StrSql & rs_Periodos_Vac!vacnro & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                Flog.writeline Espacios(Tabulador * 3) & "No Existe Periodo de Vacaciones entre " & FechaDesde & " y " & FechaHasta
                OK = False
            End If
        Case 3:    'Licencia por Nacimiento
        
        Case 4:    'Licencia por Matrimonio
        Case 5:    'Lic.Fallecimiento Familia
        Case 7:    'Licencia por Examen
        Case 8:    'Licencia por Enfermedad
        Case 9:    'Lic.por Accidente
        Case 11:   'Licencia por Maternidad
        Case 12:   'Licencia Lactancia
        Case 13:   'Lic.Accid Empresa....
        Case 14:   'Lic. Accid. ART.
        Case 15:   'Inasistencia de días
        Case 16:   'Inasist.d¡as suspensi¢n
        Case 17:   'Período de Exedencia
        Case 18:   'Inasist.reserva puesto
        Case 19:   'Lic.Mudanza
        Case 20:   '
        Case 21:   '
        Case 22:   'Lic.Gremial Interna
        Case 23:   'Lic . por Donacion de Sangre
        Case 24:   'Licencia sin goce de haberes
        Case 25:   'Desc. p/Ticket
        Case 26:   '
        Case 27:   'Llegada Tarde
        Case 28:   '
        Case 29:   'Lic.Enfermedad c / Internacion
        Case 30:   '
        Case 31:   '
        Case 32:   'Lic por Enfermedad Grave
        Case Else
            Flog.writeline Espacios(Tabulador * 3) & "Tipo de Licencia Desconocido"
            OK = False
        End Select
    Else
        ' La licencia ya existe
        OK = False
    End If
    
Exit Sub
MELocal:
    OK = False
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error Insertando Licencia. Infotipo 2001 "
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then 'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline
End Sub


Public Function Calcular_Cantidad(ByVal Cantidad As Single, ByVal Unidad As Long) As Single
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que re calcula la cantidad teniendo en cuenta la unidad de medida.
' Autor      : FGZ
' Fecha      : 25/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim aux As Single

    aux = Cantidad
    Select Case Unidad
    Case 1:    'Horas
        
    Case 2:    'Pesos
    
    Case 3:    'Dias
    
    Case 4:    'Coeficiente multiplicador
    
    Case 5:    'Cantidad
    
    Case 6:    '?
    Case 7:    '?
    Case 8:    'Porcentaje
    
    Case 9:    'Nro de Prestamo
    
    Case 10:   'Billete / Moneda
    
    Case 11:   'centimetros
    
    Case 12:   'talle
    
    Case 13:   'Entrada en grila
    
    Case 14:   'Kilometros
    
    Case 15:   'Minutos
    
    Case 16:   '
    
    Case 17:   'Kg
    
    Case 18:   'Valor
    
    Case Else
    End Select

    Calcular_Cantidad = aux
End Function


Public Sub Insertar_DDJJ(ByVal Nomina As String, ByVal Monto As Single, ByVal Fecha_Desde As String, ByVal Fecha_Hasta As String, ByVal Doc As String, ByVal RazSoc As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que mapea un cc-nomina en un itenro para un empleado
'               y lo inserta en desmen. si no encuentra el mapeo ==> no se inserta nada.
' Autor      : FGZ
' Fecha      : 28/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim i As Integer
Dim Valor As Single
Dim FechaDesde As Date
Dim FechaHasta As Date
Dim EsDesmen As Boolean
Dim Acumula As Boolean
Dim Cantidad As Boolean
Dim Prorratea As Integer

Dim rs_Item As New ADODB.Recordset
Dim rs_Desmen As New ADODB.Recordset
Dim rs_FichaRet As New ADODB.Recordset

Dim Item As Long
Dim Inserto As Boolean

On Error GoTo MELocal

FechaDesde = CDate(Fecha_Desde)
If Not EsNulo(Fecha_Hasta) Then
    FechaHasta = CDate(Fecha_Hasta)
Else
    FechaHasta = CDate("31/12/" & Year(FechaDesde))
End If

RazSoc = UCase(RazSoc)
Item = 0
Call CalcularMapeoNominaDDJJ(Nomina, Item, EsDesmen, Acumula, Cantidad)
If Acumula Then
    Cant_Acumulada = Cant_Acumulada + 1
    Monto_Acumulado = Monto_Acumulado + Monto
End If
    
If EsDesmen Then
    If Item <> 0 Then
    
        'busco la configuracion por default del item
        StrSql = "SELECT * FROM item WHERE itenro = " & Item
        If rs_Item.State = adStateOpen Then rs_Item.Close
        OpenRecordset StrSql, rs_Item
        If Not rs_Item.EOF Then
            Prorratea = rs_Item!iteprorr
        Else
            Prorratea = 0
        End If
    
        StrSql = "SELECT * FROM desmen WHERE "
        StrSql = StrSql & " itenro = " & Item
        StrSql = StrSql & " AND empleado = " & Empleado.Tercero
        StrSql = StrSql & " AND desfecdes = " & ConvFecha(FechaDesde)
        StrSql = StrSql & " AND (descuit = '" & IIf(EsNulo(Doc), " ", Doc) & "'"
        StrSql = StrSql & " AND upper(desrazsoc) = '" & IIf(EsNulo(RazSoc), " ", RazSoc) & "')"
        'StrSql = StrSql & " OR upper(desrazsoc) = '" & IIf(EsNulo(RazSoc), " ", RazSoc) & "')"
        If rs_Desmen.State = adStateOpen Then rs_Desmen.Close
        OpenRecordset StrSql, rs_Desmen
    
        If rs_Desmen.EOF Then
            Inserto = True
        Else
            Inserto = False
        End If
        
        Do While Not rs_Desmen.EOF And Not Inserto
            If rs_Desmen!desfecdes = FechaDesde Then
                If rs_Desmen!desfechas = FechaHasta Then
                    'ya existe, Actualizo
                    StrSql = "UPDATE desmen SET "
                    If Not Acumula Then
                        If Cantidad Then
                            StrSql = StrSql & " desmondec = 1 "
                        Else
                            StrSql = StrSql & " desmondec =" & Monto
                        End If
                    Else
                        If Cantidad Then
                            'StrSql = StrSql & " desmondec = desmondec + 1 "
                            StrSql = StrSql & " desmondec = " & Cant_Acumulada
                        Else
                            'StrSql = StrSql & " desmondec = desmondec + " & Monto
                            StrSql = StrSql & " desmondec = " & Monto_Acumulado
                        End If
                    End If
                    StrSql = StrSql & " ,desrazsoc ='" & RazSoc & "'"
                    StrSql = StrSql & " WHERE itenro = " & Item
                    StrSql = StrSql & " AND empleado = " & Empleado.Tercero
                    StrSql = StrSql & " AND descuit = '" & IIf(EsNulo(Doc), " ", Doc) & "'"
                Else
                    'ya existe, delimito
                    StrSql = "UPDATE desmen SET "
                    StrSql = StrSql & " desfechas =" & ConvFecha(FechaHasta)
                    StrSql = StrSql & " WHERE itenro = " & Item
                    StrSql = StrSql & " AND empleado = " & Empleado.Tercero
                    StrSql = StrSql & " AND descuit = '" & IIf(EsNulo(Doc), " ", Doc) & "'"
                End If
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                'busco el siguiente registro
            End If
            rs_Desmen.MoveNext
        Loop
        
        If Inserto Then
            StrSql = "INSERT INTO desmen ("
            StrSql = StrSql & "itenro,empleado,desmondec,desmenprorra,desano,desfecdes,desfechas,descuit,desrazsoc"
            StrSql = StrSql & ") VALUES (" & Item
            StrSql = StrSql & "," & Empleado.Tercero
            If Cantidad Then
                StrSql = StrSql & ",1"
            Else
                StrSql = StrSql & "," & Monto
            End If
            StrSql = StrSql & "," & Prorratea
            StrSql = StrSql & "," & Year(FechaDesde)
            StrSql = StrSql & "," & ConvFecha(FechaDesde)
            StrSql = StrSql & "," & ConvFecha(FechaHasta)
            StrSql = StrSql & ",'" & IIf(EsNulo(Doc), " ", Doc) & "'"
            StrSql = StrSql & ",'" & IIf(EsNulo(RazSoc), " ", RazSoc) & "'"
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
Else
    'Inserto en ficharet
    StrSql = "SELECT * FROM ficharet WHERE "
    StrSql = StrSql & " empleado = " & Empleado.Tercero
    StrSql = StrSql & " AND fecha = " & ConvFecha(FechaDesde)
    If rs_FichaRet.State = adStateOpen Then rs_FichaRet.Close
    OpenRecordset StrSql, rs_FichaRet

    If Not rs_FichaRet.EOF Then
        StrSql = "UPDATE ficharet SET "
        If Not Acumula Then
            StrSql = StrSql & " importe =" & Monto
        Else
            StrSql = StrSql & " importe = importe + " & Monto
        End If
        StrSql = StrSql & " WHERE empleado = " & Empleado.Tercero
        StrSql = StrSql & " AND fecha = " & ConvFecha(FechaDesde)
    Else
        StrSql = "INSERT INTO ficharet ("
        StrSql = StrSql & "empleado,fecha,importe,liqsistema"
        StrSql = StrSql & ") VALUES (" & Empleado.Tercero
        StrSql = StrSql & "," & ConvFecha(FechaDesde)
        StrSql = StrSql & "," & Monto
        StrSql = StrSql & ",0"
        StrSql = StrSql & " )"
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
End If


'cierro y libero
If rs_Desmen.State = adStateOpen Then rs_Desmen.Close
Set rs_Desmen = Nothing
If rs_FichaRet.State = adStateOpen Then rs_FichaRet.Close
Set rs_FichaRet = Nothing

Exit Sub
MELocal:
    'Resume Next
    Flog.writeline Espacios(Tabulador * 1) & "Error insertando DDJJ " & Err.Description
End Sub

Public Sub InsertarLogEncabezadoNovedad()
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta la linea de encabezado de novedades.
' Autor      : FGZ
' Fecha      : 07/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Aux_Linea As String

    Aux_Linea = "Novedades Cargadas"
    fNovedades.writeline Aux_Linea

    Aux_Linea = "Legajo" & Separador & "Nomina" & Separador & "Concepto" & Separador & "Parametro" & Separador & "Valor/Cantidad" & Separador & "Fecha Desde" & Separador & "Fecha Hasta"
    fNovedades.writeline Aux_Linea
    
End Sub

Public Sub InsertarLogEncabezadoCambios()
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta la linea de encabezado de Cambiso de Estructuras.
' Autor      : FGZ
' Fecha      : 07/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Aux_Linea As String

    Aux_Linea = "Cambios Generales en Estructuras"
    fCambios.writeline Aux_Linea

    Aux_Linea = "Legajo" & Separador & "Tipo Estructura" & Separador & "Estructura Anterior " & Separador & "Estructura Nueva" & Separador & "Fecha Desde" & Separador & "Fecha Hasta"
    fCambios.writeline Aux_Linea
    
End Sub


Public Sub InsertarLogNovedad(ByVal Nomina As String, ByVal concnro As Long, ByVal tpanro As Long, ByVal EsMonto As Boolean, ByVal Valor As Single, ByVal Fecha_Desde, ByVal Fecha_Hasta, ByVal Texto As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta una linea de log de la novedad insertada / actualizada.
' Autor      : FGZ
' Fecha      : 07/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Concepto As New ADODB.Recordset
Dim concepto As String
Dim Aux_Linea As String

    StrSql = "SELECT * FROM concepto "
    StrSql = StrSql & " WHERE concnro = " & concnro
    If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
    OpenRecordset StrSql, rs_Concepto
    If Not rs_Concepto.EOF Then
        concepto = rs_Concepto!Conccod
    Else
        concepto = "?" & concnro & "?"
    End If
    
    Aux_Linea = Empleado.Legajo & Separador & Nomina & Separador & concepto & Separador & tpanro & Separador & Valor & Separador & " " & Format(Fecha_Desde, "dd-mm-yyyy") & Separador & Format(Fecha_Hasta, "dd-mm-yyyy")
    fNovedades.writeline Aux_Linea
    
'cierro y libero
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
Set rs_Concepto = Nothing

End Sub



Public Sub InsertarLogcambio(ByVal Tenro As Long, ByVal EstrnroVieja As Long, ByVal EstrnroNueva As Long, ByVal Fecha_Desde, ByVal Fecha_Hasta)
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta una linea de log de la Estructura insertada / actualizada.
' Autor      : FGZ
' Fecha      : 07/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Estructura As New ADODB.Recordset
Dim rs_TEst As New ADODB.Recordset
Dim TipoEst As String
Dim EstructuraNueva As String
Dim EstructuraVieja As String
Dim Aux_Linea As String

    StrSql = "SELECT * FROM tipoestructura "
    StrSql = StrSql & " WHERE tenro = " & Tenro
    If rs_TEst.State = adStateOpen Then rs_TEst.Close
    OpenRecordset StrSql, rs_TEst
    If Not rs_TEst.EOF Then
        TipoEst = rs_TEst!tedabr
    Else
        TipoEst = "?" & Tenro & "?"
    End If

    StrSql = "SELECT * FROM estructura "
    StrSql = StrSql & " WHERE estrnro = " & EstrnroVieja
    If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
    OpenRecordset StrSql, rs_Estructura
    If Not rs_Estructura.EOF Then
        EstructuraVieja = rs_Estructura!estrdabr
    Else
        EstructuraVieja = "?" & EstrnroVieja & "?"
    End If
    
    StrSql = "SELECT * FROM estructura "
    StrSql = StrSql & " WHERE estrnro = " & EstrnroNueva
    If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
    OpenRecordset StrSql, rs_Estructura
    If Not rs_Estructura.EOF Then
        EstructuraNueva = rs_Estructura!estrdabr
    Else
        EstructuraNueva = "?" & EstrnroNueva & "?"
    End If
    
    Aux_Linea = Empleado.Legajo & Separador & TipoEst & Separador & EstructuraVieja & Separador & EstructuraNueva & Separador & " " & Format(Fecha_Desde, "dd-mm-yyyy") & Separador & Format(Fecha_Hasta, "dd-mm-yyyy")
    fCambios.writeline Aux_Linea
    
'cierro y libero
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
Set rs_Estructura = Nothing

End Sub


Public Function BuscarLegajo(ByVal Cuil As String, ByVal Aux_Legajo As String) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca un legajo con este nro de cuil.
' Autor      : FGZ
' Fecha      : 10/05/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset

    StrSql = "SELECT empleado.empleg, ter_doc.nrodoc FROM ter_doc "
    StrSql = StrSql & "INNER JOIN empleado ON empleado.ternro = ter_doc.ternro"
    StrSql = StrSql & " WHERE ter_doc.nrodoc = '" & Cuil & "'"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        BuscarLegajo = rs!empleg
    Else
        StrSql = "SELECT empleado.empleg, ter_doc.nrodoc FROM ter_doc "
        StrSql = StrSql & "INNER JOIN empleado ON empleado.ternro = ter_doc.ternro"
        StrSql = StrSql & " WHERE ter_doc.nrodoc = '" & Replace(Cuil, "-", "") & "'"
        If rs.State = adStateOpen Then rs.Close
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            BuscarLegajo = rs!empleg
        Else
            BuscarLegajo = Aux_Legajo
        End If
    End If

If rs.State = adStateOpen Then rs.Close
Set rs = Nothing
End Function


Public Function BuscarLegajo2(ByVal Aux_Legajo As String) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca un legajo con este nro de documento de tipo LSAP.
' Autor      : FGZ
' Fecha      : 10/05/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset


    If EsNulo(TipoDocLSAP) Then
        TipoDocLSAP = 28
    End If

    StrSql = "SELECT empleado.empleg, ter_doc.nrodoc FROM ter_doc "
    StrSql = StrSql & "INNER JOIN empleado ON empleado.ternro = ter_doc.ternro AND ter_doc.tidnro = " & TipoDocLSAP
    StrSql = StrSql & " WHERE ter_doc.nrodoc = '" & Aux_Legajo & "'"
    
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        BuscarLegajo2 = rs!empleg
    Else
        BuscarLegajo2 = "-1"
    End If

If rs.State = adStateOpen Then rs.Close
Set rs = Nothing
End Function


