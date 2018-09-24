Attribute VB_Name = "Module1"
Option Explicit

Public Type TDesglose
    Tenro As Long
    Estrnro_Original As Long
End Type


Public Sub Desglose_Jornada_Productiva(NroTer As Long, p_fecha As Date)

'Dim StrSql As String
Dim reghorario As Integer
Dim categoria As Integer
Dim producto As Integer
Dim PropJornalProduccion As Single
Dim l_achdnro As Long
Dim auxi As String
Dim rs As New ADODB.Recordset
Dim HorasProduccion As Integer
Dim HorasProdDia As Integer
Dim HorasTrabajadas As Single
Dim HorasTurno As Single

Dim horres As Single

'Borro para reprocesar
StrSql = "delete from gti_achdiario_estr where achdnro in(" & _
"SELECT achdnro FROM gti_achdiario WHERE ternro = " & NroTer & _
" AND achdfecha = " & ConvFecha(p_fecha) & _
" AND achdmanual = 0)"
objConn.Execute StrSql, , adExecuteNoRecords

StrSql = "delete from gti_achdiario where ternro = " & NroTer & _
" AND achdfecha = " & ConvFecha(p_fecha) & " AND achdmanual = 0"
objConn.Execute StrSql, , adExecuteNoRecords

'Busco las horas Produccion
StrSql = "SELECT confval FROM confrep WHERE repnro = 53 AND" & _
" conftipo = 'TH' AND confnrocol = 4"
OpenRecordset StrSql, rs
If rs.EOF Then
    Exit Sub
End If
HorasProduccion = rs!confval

'Busco el total de horas producción del día
StrSql = "SELECT * FROM gti_acumdiario " & _
    " WHERE (ternro = " & NroTer & ") AND " & _
    " (adfecha = " & ConvFecha(p_fecha) & ")" & _
    " AND thnro = " & HorasProduccion
OpenRecordset StrSql, rs
If rs.EOF Then
    Exit Sub
End If

HorasTrabajadas = rs!adcanthoras

'Busco las horas a desglosar en el confrep
StrSql = "SELECT confval FROM confrep WHERE repnro = 53 AND" & _
" conftipo = 'TH' AND confnrocol = 5"
OpenRecordset StrSql, rs
If rs.EOF Then
  Exit Sub
End If
HorasProdDia = rs!confval

' Busco en el histórico las estructuras
StrSql = " SELECT estrnro FROM his_estructura"
StrSql = StrSql & " WHERE his_estructura.tenro = 3 "
StrSql = StrSql & " AND his_estructura.ternro = " & NroTer
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(p_fecha)
StrSql = StrSql & " AND ((" & ConvFecha(p_fecha) & " <= htethasta) OR (htethasta is null))"
OpenRecordset StrSql, rs
If rs.EOF Then
  Exit Sub
End If
categoria = rs!estrnro

StrSql = " SELECT estrnro FROM his_estructura"
StrSql = StrSql & " WHERE his_estructura.tenro = 21 "
StrSql = StrSql & " AND his_estructura.ternro = " & NroTer
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(p_fecha)
StrSql = StrSql & " AND ((" & ConvFecha(p_fecha) & " <= htethasta) OR (htethasta is null))"
OpenRecordset StrSql, rs
If rs.EOF Then
  Exit Sub
End If
reghorario = rs!estrnro

StrSql = " SELECT estrnro FROM his_estructura"
StrSql = StrSql & " WHERE his_estructura.tenro = 38 "
StrSql = StrSql & " AND his_estructura.ternro = " & NroTer
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(p_fecha)
StrSql = StrSql & " AND ((" & ConvFecha(p_fecha) & " <= htethasta) OR (htethasta is null))"
OpenRecordset StrSql, rs
If rs.EOF Then
  Exit Sub
End If
producto = rs!estrnro


'Busco las horas producción del día
StrSql = "SELECT * FROM gti_acumdiario " & _
    " WHERE (ternro = " & NroTer & ") AND " & _
    " (adfecha = " & ConvFecha(p_fecha) & ")" & _
    " AND thnro = " & HorasProdDia
OpenRecordset StrSql, rs
   
If Not rs.EOF Then
    
    If (rs!adcanthoras = 0) Or (HorasTrabajadas = 0) Then Exit Sub
        
    HorasTurno = HorasTrabajadas / rs!adcanthoras
    PropJornalProduccion = Round(HorasTrabajadas / HorasTurno, 2)
    
    'Busco los partes de movilidad del día
    StrSql = " SELECT DISTINCT ternro, gmdnro, gmdhoras FROM gti_movdet "
    StrSql = StrSql & " WHERE gmdfecdesde <= " & ConvFecha(p_fecha)
    StrSql = StrSql & " AND " & ConvFecha(p_fecha) & " <= gmdfechasta "
    StrSql = StrSql & " AND ternro = " & NroTer
    OpenRecordset StrSql, objRs
    
    ' Si existen entonces
    Do While Not objRs.EOF
        'Hago el desgloce en las estructuras especificadas en el parte
        'de movilidad, por el tiempo definido en el.
    
        PropJornalProduccion = PropJornalProduccion - Round(objRs!gmdhoras / HorasTurno, 2)
        
        ' Inserto en la tabla de desglose con horas producción de jornada
        StrSql = "INSERT INTO gti_achdiario "
        StrSql = StrSql & "(achdcanthoras, achdfecha, achdmanual, achdvalido,ternro,thnro) "
        StrSql = StrSql & " VALUES (" & Round(objRs!gmdhoras / HorasTurno, 2) & ","
        StrSql = StrSql & ConvFecha(p_fecha) & ",0,-1,"
        StrSql = StrSql & NroTer & ", " & HorasProdDia & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'l_achdnro = getLastIdentity(objConn, "GTI_ACHDIARIO")
                
        StrSql = "select max(achdnro) as next_id from GTI_ACHDIARIO"
        OpenRecordset StrSql, rs
        l_achdnro = rs!next_id
        
        'Dentro del parte, busco una referencia a la estructura empaque
        StrSql = "SELECT * FROM gti_movdet_estr WHERE gmdnro = " & objRs!gmdnro
        StrSql = StrSql & " AND tenro = 21"
        OpenRecordset StrSql, rs

        StrSql = "INSERT INTO gti_achdiario_estr "
        StrSql = StrSql & "(achdnro, tenro, estrnro, achdfecha) "
        StrSql = StrSql & " VALUES (" & l_achdnro & ",21"
        'Reemplazo el empaque del empleado por el del parte
        If Not rs.EOF Then
            StrSql = StrSql & ", " & rs!estrnro & ","
        Else
            StrSql = StrSql & ", " & reghorario & ","
        End If
        StrSql = StrSql & ConvFecha(p_fecha) & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'Dentro del parte, busco una referencia a la estructura producto
        StrSql = "SELECT * FROM gti_movdet_estr WHERE gmdnro = " & objRs!gmdnro
        StrSql = StrSql & " AND tenro = 38"
        OpenRecordset StrSql, rs
        
'        ' FGZ - 01/03/2004
'        If rs.EOF Then
'            StrSql = "INSERT INTO gti_achdiario_estr "
'            StrSql = StrSql & "(achdnro, tenro, estrnro, achdfecha)"
'            StrSql = StrSql & " VALUES (" & l_achdnro & ",38"
'            StrSql = StrSql & "," & producto & ","
'            StrSql = StrSql & ConvFecha(p_fecha) & ")"
'            objConn.Execute StrSql, , adExecuteNoRecords
'        End If
'        Do While Not rs.EOF
'            StrSql = "INSERT INTO gti_achdiario_estr "
'            StrSql = StrSql & "(achdnro, tenro, estrnro, achdfecha)"
'            StrSql = StrSql & " VALUES (" & l_achdnro & ",38"
'            'Reemplazo el producto del empleado por el del parte
'            StrSql = StrSql & "," & rs!estrnro & ","
'            StrSql = StrSql & ConvFecha(p_fecha) & ")"
'            objConn.Execute StrSql, , adExecuteNoRecords
'
'            rs.MoveNext
'        Loop
        
        ' antes del 01/03/2004
        StrSql = "INSERT INTO gti_achdiario_estr "
        StrSql = StrSql & "(achdnro, tenro, estrnro, achdfecha)"
        StrSql = StrSql & " VALUES (" & l_achdnro & ",38"
        'Reemplazo el producto del empleado por el del parte
        If Not rs.EOF Then
            StrSql = StrSql & "," & rs!estrnro & ","
        Else
            StrSql = StrSql & "," & producto & ","
        End If
        StrSql = StrSql & ConvFecha(p_fecha) & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        ' antes del 01/03/2004
        
        
        'Dentro del parte, busco una referencia a la estructura categoria
        StrSql = "SELECT * FROM gti_movdet_estr WHERE gmdnro = " & objRs!gmdnro
        StrSql = StrSql & " AND tenro = 3"
        OpenRecordset StrSql, rs
        
        StrSql = "INSERT INTO gti_achdiario_estr "
        StrSql = StrSql & "(achdnro, tenro, estrnro, achdfecha) "
        StrSql = StrSql & " VALUES (" & l_achdnro & ",3"
        'Reemplazo la categoria del empleado por el del parte
        If Not rs.EOF Then
            StrSql = StrSql & ", " & rs!estrnro & ","
        Else
            StrSql = StrSql & ", " & categoria & ","
        End If
        StrSql = StrSql & ConvFecha(p_fecha) & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        objRs.MoveNext
    Loop
    
    ' si aún quedan horas por desglosar, entonces uso las estructuras
    ' por descarte
    If PropJornalProduccion > 0 Then
        
        ' Inserto en la tabla de desglose
        StrSql = "INSERT INTO gti_achdiario "
        StrSql = StrSql & "(achdcanthoras, achdfecha, achdmanual, achdvalido,ternro,thnro) "
        StrSql = StrSql & " VALUES (" & PropJornalProduccion & ","
        StrSql = StrSql & ConvFecha(p_fecha) & ",0,-1,"
        StrSql = StrSql & NroTer & "," & HorasProdDia & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
               
        l_achdnro = getLastIdentity(objConn, "GTI_ACHDIARIO")
                
        StrSql = "select max(achdnro) as next_id from GTI_ACHDIARIO"
        OpenRecordset StrSql, rs
        l_achdnro = rs!next_id
        
        'Inserto en la tabla de desgloce por estructura
        '(un registro por cada estructura)
        
        StrSql = "INSERT INTO gti_achdiario_estr "
        StrSql = StrSql & "(achdnro, tenro, estrnro, achdfecha) "
        StrSql = StrSql & " VALUES (" & l_achdnro & ",3"
        StrSql = StrSql & ", " & categoria & ","
        StrSql = StrSql & ConvFecha(p_fecha) & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
   
        StrSql = "INSERT INTO gti_achdiario_estr "
        StrSql = StrSql & "(achdnro, tenro, estrnro, achdfecha) "
        StrSql = StrSql & " VALUES (" & l_achdnro & ",21"
        StrSql = StrSql & ", " & reghorario & ","
        StrSql = StrSql & ConvFecha(p_fecha) & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    
        StrSql = "INSERT INTO gti_achdiario_estr "
        StrSql = StrSql & "(achdnro, tenro, estrnro, achdfecha) "
        StrSql = StrSql & " VALUES (" & l_achdnro & ",38"
        StrSql = StrSql & ", " & producto & ","
        StrSql = StrSql & ConvFecha(p_fecha) & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        rs.Close
    End If
    
End If

If objRs.State = adStateOpen Then objRs.Close
If rs.State = adStateOpen Then rs.Close

Set rs = Nothing

End Sub

Public Sub Desglose_RelevosV2(ByVal NroTer As Long, ByVal Fecha As Date)
' ------------------------------------------------------------------
' Descripcion: Procedimiento que genera los desgloces de horas de los partes de relevos de estructura. Esta versión genera desgloce de multiples estructuras a diferencia de la versión 1
' Autor: EAM - 14/05/2013
' Ultima modificacion:
' ------------------------------------------------------------------
Dim i As Long
Dim Z As Long
Dim Aux_Relnro As Long
Dim tipoEstructura As Long
Dim TotalHS As Single

Dim rs_AD As New ADODB.Recordset
Dim rs_Estr As New ADODB.Recordset
Dim rs_Rel As New ADODB.Recordset
Dim rs_RelDet As New ADODB.Recordset

'Borro la tabla de desgloce diario para reprocesar
StrSql = "DELETE FROM gti_desgldiario WHERE ternro = " & NroTer & _
        " AND fecha = " & ConvFecha(Fecha) & " AND manual = 0"
objConn.Execute StrSql, , adExecuteNoRecords


'Busco el total de horas producción del día
StrSql = "SELECT * FROM gti_acumdiario " & _
        " INNER JOIN tiphora ON gti_acumdiario.thnro = tiphora.thnro AND tiphora.thprod = -1" & _
        " WHERE (ternro = " & NroTer & ") AND (adfecha = " & ConvFecha(Fecha) & ")"
OpenRecordset StrSql, rs_AD

If rs_AD.EOF Then
    If depurar Then
        Flog.writeline Espacios(Tabulador * 3) & "No hay Horas produccion para desglosar."
    End If
Else
    'EAM (v5.32)- Busco si hay partes de relevos
    StrSql = "SELECT distinct gti_relevo_estruc.relestnro,gti_relevo_estruc.estrnro,gti_relevo_estruc.tenro FROM gti_relevo_estruc " & _
            " INNER JOIN gti_relevo_estruc_det ON gti_relevo_estruc.relestnro = gti_relevo_estruc_det.relestnro " & _
            " WHERE gti_relevo_estruc.ternro = " & NroTer & _
            " AND (gti_relevo_estruc.relestfecdesde <= " & ConvFecha(Fecha) & " AND gti_relevo_estruc.relestfechasta >= " & ConvFecha(Fecha) & ")" & _
            " AND gti_relevo_estruc_det.fecha = " & ConvFecha(Fecha) & " ORDER BY gti_relevo_estruc.tenro"
    OpenRecordset StrSql, rs_Rel
    If Not rs_Rel.EOF Then
    Do While Not rs_Rel.EOF
        Aux_Relnro = rs_Rel!relestnro
        tipoEstructura = rs_Rel!Tenro
        'Estructura activa del empleado
        
        Do While Not rs_AD.EOF
            TotalHS = 0
            
            'EAM (v5.32)- Busco el detalle de horas desglozadas
            StrSql = "SELECT canths,gti_relevo_estruc_det_hs.estrnro FROM gti_relevo_estruc_det " & _
                    " INNER JOIN  gti_relevo_estruc_det_hs ON gti_relevo_estruc_det_hs.relestdetnro = gti_relevo_estruc_det.relestdetnro " & _
                    " WHERE gti_relevo_estruc_det.fecha=" & ConvFecha(Fecha) & " AND gti_relevo_estruc_det_hs.thnro= " & rs_AD!thnro & _
                    " AND gti_relevo_estruc_det.relestnro = " & rs_Rel!relestnro
            OpenRecordset StrSql, rs_RelDet
        
                Do While Not rs_RelDet.EOF
                    'EAM (v5.32) - Inserto el desglose
                    StrSql = "INSERT INTO gti_desgldiario " & _
                            " (ternro, fecha, thnro, canthoras, horas, manual, valido, te1, estrnro1, te2, estrnro2) " & _
                            " VALUES ( " & NroTer & "," & ConvFecha(Fecha) & "," & rs_AD!thnro & "," & (CDbl(rs_RelDet!canths)) & _
                            "," & CHoras(rs_RelDet!canths, 60) & ",0" & ",-1," & rs_Rel!Tenro & "," & rs_Rel!estrnro & "," & rs_Rel!Tenro & "," & rs_RelDet!estrnro & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    TotalHS = CDbl(TotalHS) + CDbl(rs_RelDet!canths)
                    rs_RelDet.MoveNext
                Loop
        
            If (CDbl(rs_AD!adcanthoras) > CDbl(TotalHS)) Then
                StrSql = "INSERT INTO gti_desgldiario " & _
                        " (ternro, fecha, thnro, canthoras, horas, manual, valido, te1, estrnro1, te2, estrnro2) " & _
                        " VALUES ( " & NroTer & "," & ConvFecha(Fecha) & "," & rs_AD!thnro & "," & (CDbl(rs_AD!adcanthoras) - CDbl(TotalHS)) & _
                        "," & CHoras(CDbl(rs_AD!adcanthoras) - CDbl(TotalHS), 60) & ",0" & ",-1," & rs_Rel!Tenro & "," & rs_Rel!estrnro & "," & rs_Rel!Tenro & "," & rs_Rel!estrnro & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                Flog.writeline Espacios(Tabulador * 3) & "La cantidad de horas de relevos de estructuras es mayor al total de horas."
            End If
                                
            rs_AD.MoveNext
        Loop
        rs_Rel.MoveNext
    Loop
    Else
        Flog.writeline Espacios(Tabulador * 3) & "No existen partes de relevos."
    End If

End If



        

        
'Cierro y libero
If rs_AD.State = adStateOpen Then rs_AD.Close
Set rs_AD = Nothing

If rs_Estr.State = adStateOpen Then rs_Estr.Close
Set rs_Estr = Nothing

If rs_Rel.State = adStateOpen Then rs_Rel.Close
Set rs_Rel = Nothing



End Sub

Public Sub Desglose_Relevos(ByVal NroTer As Long, ByVal Fecha As Date)
' ------------------------------------------------------------------
' Descripcion: Procedimiento que genera los desgloces a partir de
'               los partes diarios de relevos.
' Autor: FGZ - 07/10/2010
' Ultima modificacion:
' ------------------------------------------------------------------
'gti_desgldiario(
'    desgnro int IDENTITY(1,1) NOT NULL,
'    canthoras decimal (5, 2) NULL,
'    fecha datetime NOT NULL,
'    manual smallint NOT NULL,
'    valido smallint NOT NULL,
'    ternro int NOT NULL,
'    thnro int NOT NULL,
'    horas varchar(10) NULL,
'    te1 int NOT NULL DEFAULT 0,
'    estrnro1 int NOT NULL DEFAULT 0,
'    te2 int NOT NULL DEFAULT 0,
'    estrnro2 int NOT NULL DEFAULT 0,
'    te3 int NOT NULL DEFAULT 0,
'    estrnro3 int NOT NULL DEFAULT 0,
'    te4 int NOT NULL DEFAULT 0,
'    estrnro4 int NOT NULL DEFAULT 0,
'    te5 int NOT NULL DEFAULT 0,
'    estrnro5 int NOT NULL DEFAULT 0
')
' ------------------------------------------------------------------
Dim i As Long
Dim Z As Long
Dim Aux_Relnro As Long

Dim rs_AD As New ADODB.Recordset
Dim rs_Estr As New ADODB.Recordset
Dim rs_Rel As New ADODB.Recordset
Dim rs_RelDet As New ADODB.Recordset

Dim Estr_Def(1 To 5) As TEstr
Dim Estr_Ind(1 To 5) As Long
Dim Lista_TE As Collection

'Borro para reprocesar
StrSql = "DELETE FROM gti_desgldiario WHERE ternro = " & NroTer
StrSql = StrSql & " AND fecha = " & ConvFecha(Fecha)
StrSql = StrSql & " AND manual = 0"
objConn.Execute StrSql, , adExecuteNoRecords



'Busco el total de horas producción del día
StrSql = "SELECT * FROM gti_acumdiario "
StrSql = StrSql & " INNER JOIN tiphora ON gti_acumdiario.thnro = tiphora.thnro AND tiphora.thprod = -1"
StrSql = StrSql & " WHERE (ternro = " & NroTer & ") AND "
StrSql = StrSql & " (adfecha = " & ConvFecha(Fecha) & ")"
OpenRecordset StrSql, rs_AD
If rs_AD.EOF Then
    If depurar Then
        Flog.writeline Espacios(Tabulador * 3) & "No hay Horas produccion para desglosar."
    End If
Else
    'Busco las estrucutras default (asignadas por historico de estr) para los tipos de estructuras con alcance
    StrSql = "SELECT * FROM his_estructura "
    StrSql = StrSql & " INNER JOIN alcance_testr ON his_estructura.tenro = alcance_testr.tenro AND alcance_testr.tanro = 22"
    StrSql = StrSql & " WHERE ternro = " & NroTer
    StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(Fecha) & ")"
    StrSql = StrSql & " AND ((" & ConvFecha(Fecha) & " <= his_estructura.htethasta) OR (his_estructura.htethasta IS NULL))"
    StrSql = StrSql & " ORDER BY alcance_testr.alteorden, his_estructura.tenro"
    OpenRecordset StrSql, rs_Estr
    i = 1
    Do While Not rs_Estr.EOF And i <= 5
        Estr_Def(i).Tenro = rs_Estr!Tenro
        Estr_Def(i).estrnro = rs_Estr!estrnro
        Estr_Def(i).rel = rs_Estr!estrnro
        Estr_Def(i).relCantHs = 0
        Estr_Ind(i) = rs_Estr!Tenro
        
        i = i + 1
        rs_Estr.MoveNext
    Loop
        
    'Busco si hay partes de relevos ----------------------------------------------
    'gti_relevos(
    'relnro int IDENTITY(1,1) NOT NULL,
    'gcpnro int NOT NULL,
    'relfecdesde datetime NULL,
    'relfechasta datetime NULL,
    'ternro int NOT NULL,
    'tenro int NOT NULL,
    'estrnro int NOT NULL
    ')
    '
    'gti_relevos_det(
    'reldetnro int IDENTITY(1,1) NOT NULL,
    'relnro int NOT NULL,
    'fecha datetime NULL
    ')

    StrSql = "SELECT * FROM gti_relevos "
    StrSql = StrSql & " INNER JOIN gti_relevos_det ON gti_relevos.relnro = gti_relevos_det.relnro "
    StrSql = StrSql & " WHERE gti_relevos.ternro = " & NroTer
    StrSql = StrSql & " AND (gti_relevos.relfecdesde <= " & ConvFecha(Fecha) & " AND gti_relevos.relfechasta >= " & ConvFecha(Fecha) & ")"
    StrSql = StrSql & " AND gti_relevos_det.fecha = " & ConvFecha(Fecha)
    StrSql = StrSql & " ORDER BY gti_relevos.tenro"
    OpenRecordset StrSql, rs_Rel
    Do While Not rs_Rel.EOF
        Aux_Relnro = rs_Rel!relnro
        Z = IndiceTE(rs_Rel!Tenro, Estr_Ind)
        If Z > 0 Then
            Estr_Def(Z).rel = rs_Rel!estrnro
        End If

        rs_Rel.MoveNext
    Loop
    
    'Inserto el detalle
    Do While Not rs_AD.EOF
        

        StrSql = "SELECT canths FROM gti_relevos_det " & _
                "INNER JOIN  gti_relevos_det_hs on gti_relevos_det_hs.reldetnro= gti_relevos_det.reldetnro " & _
                "WHERE gti_relevos_det.fecha=" & ConvFecha(Fecha) & " AND gti_relevos_det_hs.thnro= " & rs_AD!thnro & _
                " AND gti_relevos_det.relnro = " & Aux_Relnro
        OpenRecordset StrSql, rs_RelDet
        
        
        
        If rs_RelDet.EOF Then
            'INSERT
            StrSql = "INSERT INTO gti_desgldiario "
            StrSql = StrSql & "(ternro, fecha, thnro, canthoras, horas, manual, valido"
            For i = 1 To 5
                If Estr_Def(i).Tenro <> 0 Then
                    StrSql = StrSql & ", te" & i & ", estrnro" & i
                End If
            Next i
            StrSql = StrSql & ") VALUES ("
            StrSql = StrSql & NroTer
            StrSql = StrSql & "," & ConvFecha(Fecha)
            StrSql = StrSql & "," & rs_AD!thnro
            StrSql = StrSql & "," & rs_AD!adcanthoras
            StrSql = StrSql & ",'" & rs_AD!Horas & "'"
            StrSql = StrSql & ",0"
            StrSql = StrSql & ",-1"
            For i = 1 To 5
                If Estr_Def(i).Tenro <> 0 Then
                    StrSql = StrSql & "," & Estr_Def(i).Tenro
                    StrSql = StrSql & "," & Estr_Def(i).rel
                End If
            Next i
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            'INSERT
            StrSql = "INSERT INTO gti_desgldiario "
            StrSql = StrSql & "(ternro, fecha, thnro, canthoras, horas, manual, valido"
            For i = 1 To 5
                If Estr_Def(i).Tenro <> 0 Then
                    StrSql = StrSql & ", te" & i & ", estrnro" & i
                End If
            Next i
            StrSql = StrSql & ") VALUES ("
            StrSql = StrSql & NroTer
            StrSql = StrSql & "," & ConvFecha(Fecha)
            StrSql = StrSql & "," & rs_AD!thnro
            StrSql = StrSql & "," & (CDbl(rs_AD!adcanthoras) - CDbl(rs_RelDet!canths))
            StrSql = StrSql & ",'" & rs_AD!Horas & "'"
            StrSql = StrSql & ",0"
            StrSql = StrSql & ",-1"
            For i = 1 To 5
                If Estr_Def(i).Tenro <> 0 Then
                    StrSql = StrSql & "," & Estr_Def(i).Tenro
                    StrSql = StrSql & "," & Estr_Def(i).estrnro
                End If
            Next i
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            
            If rs_RelDet!canths > 0 Then
                'INSERT
                StrSql = "INSERT INTO gti_desgldiario "
                StrSql = StrSql & "(ternro, fecha, thnro, canthoras, horas, manual, valido"
                For i = 1 To 5
                    If Estr_Def(i).Tenro <> 0 Then
                        StrSql = StrSql & ", te" & i & ", estrnro" & i
                    End If
                Next i
                StrSql = StrSql & ") VALUES ("
                StrSql = StrSql & NroTer
                StrSql = StrSql & "," & ConvFecha(Fecha)
                StrSql = StrSql & "," & rs_AD!thnro
                StrSql = StrSql & "," & rs_RelDet!canths
                StrSql = StrSql & ",'" & rs_AD!Horas & "'"
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",-1"
                For i = 1 To 5
                    If Estr_Def(i).Tenro <> 0 Then
                        StrSql = StrSql & "," & Estr_Def(i).Tenro
                        StrSql = StrSql & "," & Estr_Def(i).rel
                    End If
                Next i
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
    
        rs_AD.MoveNext
    Loop
End If



        

        
'Cierro y libero
If rs_AD.State = adStateOpen Then rs_AD.Close
Set rs_AD = Nothing

If rs_Estr.State = adStateOpen Then rs_Estr.Close
Set rs_Estr = Nothing

If rs_Rel.State = adStateOpen Then rs_Rel.Close
Set rs_Rel = Nothing



End Sub



Public Sub Desglose_Jornada_Productiva_Nuevo(ByVal NroTer As Long, ByVal p_fecha As Date)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de genera el desglose de Jornada productiva.
' Autor      : FGZ
' Fecha      : 17/11/2005
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim Estructura(1 To 5) As TDesglose
Dim Usa_TE(1 To 5) As Boolean
Dim CantidadTE As Integer
Dim i As Integer

Dim THorasProduccion As Long
Dim THorasProdDia As Long
Dim HorasTrabajadas As Single
Dim HorasTrabajadasDia As Single

Dim PropJornalProduccion As Single
Dim l_achdnro As Long
Dim HorasTurno As Single

Dim rs As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Partes As New ADODB.Recordset
Dim rs_Partes_Estr As New ADODB.Recordset
Dim Seguir As Boolean
Dim TotHorHHMM As String
Dim rs_FT As New ADODB.Recordset

If depurar Then
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Desglose_Jornada_Productiva"
End If

'Borro para reprocesar
StrSql = "DELETE FROM gti_achdiario_estr "
StrSql = StrSql & " WHERE achdnro IN( "
    StrSql = StrSql & " SELECT achdnro FROM gti_achdiario "
    StrSql = StrSql & " WHERE (ternro = " & NroTer
    StrSql = StrSql & " AND achdfecha = " & ConvFecha(p_fecha)
    StrSql = StrSql & " AND achdmanual = 0) "
StrSql = StrSql & ")"
objConn.Execute StrSql, , adExecuteNoRecords

StrSql = "DELETE FROM gti_achdiario "
StrSql = StrSql & " WHERE ternro = " & NroTer
StrSql = StrSql & " AND achdfecha = " & ConvFecha(p_fecha)
StrSql = StrSql & " AND achdmanual = 0"
objConn.Execute StrSql, , adExecuteNoRecords

'Busco las horas Produccion
StrSql = "SELECT confval FROM confrep "
StrSql = StrSql & " WHERE repnro = 53 "
StrSql = StrSql & " AND conftipo = 'TH' AND confnrocol = 4"
If rs.State = adStateOpen Then rs.Close
OpenRecordset StrSql, rs
If rs.EOF Then
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No esta configurado el tipo de hora produccion. Confrep 53 columna 4."
    End If
    Exit Sub
End If
THorasProduccion = rs!confval

'Busco el total de horas producción del día
StrSql = "SELECT adcanthoras FROM gti_acumdiario "
StrSql = StrSql & " WHERE ternro = " & NroTer
StrSql = StrSql & " AND adfecha = " & ConvFecha(p_fecha)
StrSql = StrSql & " AND thnro = " & THorasProduccion
If rs.State = adStateOpen Then rs.Close
OpenRecordset StrSql, rs
If rs.EOF Then
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No hay tipo de hora " & THorasProduccion & " para la fecha " & p_fecha
    End If
    Exit Sub
End If
HorasTrabajadas = rs!adcanthoras

'Busco las horas a desglosar en el confrep
StrSql = "SELECT confval FROM confrep WHERE repnro = 53 "
StrSql = StrSql & " AND conftipo = 'TH' AND confnrocol = 5"
If rs.State = adStateOpen Then rs.Close
OpenRecordset StrSql, rs
If rs.EOF Then
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No esta configurado el tipo de hora a desglozar. Confrep 53 columna 5."
    End If
    Exit Sub
End If
THorasProdDia = rs!confval

'Busco los tipos de estructuras a Desglozar
StrSql = "SELECT confval, confnrocol FROM confrep WHERE repnro = 53 "
StrSql = StrSql & " AND confnrocol >= 50 AND confnrocol <= 54"
StrSql = StrSql & " ORDER BY confnrocol"
If rs.State = adStateOpen Then rs.Close
OpenRecordset StrSql, rs
If rs.EOF Then
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No hay configurado ningun tipo de estructura a desglozar."
    End If
    Exit Sub
End If

'Inicializo las estructuras del desglose
For i = 1 To 5
    Estructura(i).Tenro = 0
    Estructura(i).Estrnro_Original = 0
    Usa_TE(i) = False
Next i

i = 0
Do While Not rs.EOF And i <= 5
    If Not EsNulo(rs!confval) Then
        i = i + 1
        Estructura(i).Tenro = rs!confval
        Usa_TE(i) = True
    
        'Busco en el histórico las estructuras
        StrSql = " SELECT estrnro FROM his_estructura"
        StrSql = StrSql & " WHERE his_estructura.tenro = " & Estructura(i).Tenro
        StrSql = StrSql & " AND his_estructura.ternro = " & NroTer
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(p_fecha)
        StrSql = StrSql & " AND ((" & ConvFecha(p_fecha) & " <= htethasta) OR (htethasta is null))"
        If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
        OpenRecordset StrSql, rs_Estructura
        If rs_Estructura.EOF Then
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "El empleado no tiene activa ningun estructura de tipo " & Estructura(i).Tenro & "."
            End If
            'Exit Sub
        Else
            Estructura(i).Estrnro_Original = rs_Estructura!estrnro
        End If
    End If

    rs.MoveNext
Loop
CantidadTE = i

'FGZ - 10/07/2008 - Le agregué este control porque sino insertaba registros con estrnro = 0
Seguir = True
For i = 1 To 5
    If Estructura(i).Tenro <> 0 Then
        If Estructura(i).Estrnro_Original = 0 Then
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "El empleado no tiene activa ningun estructura de tipo " & Estructura(i).Tenro & "."
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No se puede hacer ningun desgloce."
            End If
            Seguir = False
        End If
    End If
Next i
If Not Seguir Then
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "El empleado no tiene activa alguna estructura necesaria para el desgloce."
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No se efectuará ningun desgloce."
    End If
    GoTo FIN
End If


'Busco las horas producción del día
StrSql = "SELECT adcanthoras FROM gti_acumdiario "
StrSql = StrSql & " WHERE ternro = " & NroTer
StrSql = StrSql & " AND adfecha = " & ConvFecha(p_fecha)
StrSql = StrSql & " AND thnro = " & THorasProdDia
If rs.State = adStateOpen Then rs.Close
OpenRecordset StrSql, rs
If rs.EOF Then
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "El empleado no tiene horas produccion (" & THorasProdDia & ") en la fecha " & p_fecha
    End If
    'esto es idea mia, poner cantidad = 1 por default
    'HorasTrabajadasDia = 1
    Exit Sub
Else
    HorasTrabajadasDia = rs!adcanthoras
End If
If (HorasTrabajadasDia = 0) Or (HorasTrabajadas = 0) Then
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "La cantidad de horas 0."
    End If
    Exit Sub
End If
    
HorasTurno = HorasTrabajadas / HorasTrabajadasDia
PropJornalProduccion = Round(HorasTrabajadas / HorasTurno, 2)

'Busco los partes de movilidad del día
'StrSql = " SELECT DISTINCT ternro, gmdnro, gmdhoras FROM gti_movdet "
'StrSql = StrSql & " WHERE gmdfecdesde <= " & ConvFecha(p_fecha)
'StrSql = StrSql & " AND " & ConvFecha(p_fecha) & " <= gmdfechasta "
'StrSql = StrSql & " AND ternro = " & NroTer
'If rs_Partes.State = adStateOpen Then rs_Partes.Close



StrSql = " SELECT DISTINCT gti_movdet.gcpnro,gti_movdet.ternro, gti_movdet.gmdnro, gti_movdet.gmdhoras, gti_cabparte.ft, gti_cabparte.ftap FROM gti_cabparte "
StrSql = StrSql & " INNER JOIN gti_movdet ON gti_cabparte.gcpnro = gti_movdet.gcpnro "
StrSql = StrSql & " WHERE gmdfecdesde <= " & ConvFecha(p_fecha)
StrSql = StrSql & " AND " & ConvFecha(p_fecha) & " <= gmdfechasta "
StrSql = StrSql & " AND ternro = " & NroTer
StrSql = StrSql & " AND (ft = 0 OR (ft = -1 AND ftap = -1))"
OpenRecordset StrSql, rs_Partes
Do While Not rs_Partes.EOF
    'Hago el desgloce en las estructuras especificadas en el parte de movilidad, por el tiempo definido en el.
    
    
    StrSql = "SELECT input_ft.idnro,input_ft.origen, gti_cabparte.ft, gti_cabparte.ftap FROM input_ft "
    StrSql = StrSql & " INNER JOIN gti_cabparte ON input_ft.origen = gti_cabparte.gcpnro "
    StrSql = StrSql & " WHERE idtipoinput = 9 "
    StrSql = StrSql & " AND origen = " & rs_Partes!GCPNRO
    OpenRecordset StrSql, rs_FT
    If Not rs_FT.EOF Then
        If rs_FT!ftap = -1 Then
            Call InsertarFT(rs_FT!idnro, 9, rs_FT!Origen)
        End If
    End If
           
    PropJornalProduccion = PropJornalProduccion - Round(rs_Partes!gmdhoras / HorasTurno, 2)
    
    'Inserto en la tabla de desglose con horas producción de jornada
    TotHorHHMM = CHoras(Round(rs_Partes!gmdhoras / HorasTurno, 2), 60)
    
    StrSql = "INSERT INTO gti_achdiario "
    StrSql = StrSql & "(horas, achdcanthoras, achdfecha, achdmanual, achdvalido,ternro,thnro) "
    StrSql = StrSql & " VALUES (" & TotHorHHMM & "," & Round(rs_Partes!gmdhoras / HorasTurno, 2) & ","
    StrSql = StrSql & ConvFecha(p_fecha) & ",0,-1,"
    StrSql = StrSql & NroTer & ", " & THorasProdDia & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    l_achdnro = getLastIdentity(objConn, "gti_achdiario")
            
    'busco dentro del parte cada TE a desglozar
    For i = 1 To CantidadTE
        If Usa_TE(i) Then
            StrSql = "SELECT * FROM gti_movdet_estr WHERE gmdnro = " & rs_Partes!gmdnro
            StrSql = StrSql & " AND tenro = " & Estructura(i).Tenro
            If rs_Partes_Estr.State = adStateOpen Then rs_Partes_Estr.Close
            OpenRecordset StrSql, rs_Partes_Estr
            
            StrSql = "INSERT INTO gti_achdiario_estr "
            StrSql = StrSql & "(achdnro, tenro, estrnro, achdfecha) "
            StrSql = StrSql & " VALUES ("
            StrSql = StrSql & l_achdnro & ","
            StrSql = StrSql & Estructura(i).Tenro & ","
            If Not rs_Partes_Estr.EOF Then
                StrSql = StrSql & rs_Partes_Estr!estrnro & ","
            Else
                'Reemplazo el empaque del empleado por el del parte
                StrSql = StrSql & Estructura(i).Estrnro_Original & ","
            End If
            StrSql = StrSql & ConvFecha(p_fecha) & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    Next i
    
    rs_Partes.MoveNext
Loop


' si aún quedan horas por desglosar, entonces uso las estructuras
' por descarte
If PropJornalProduccion > 0 Then
    ' Inserto en la tabla de desglose
    TotHorHHMM = CHoras(PropJornalProduccion, 60)
    
    StrSql = "INSERT INTO gti_achdiario "
    StrSql = StrSql & "(horas, achdcanthoras, achdfecha, achdmanual, achdvalido,ternro,thnro) "
    StrSql = StrSql & " VALUES (" & TotHorHHMM & "," & PropJornalProduccion & ","
    StrSql = StrSql & ConvFecha(p_fecha) & ",0,-1,"
    StrSql = StrSql & NroTer & "," & THorasProdDia & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    l_achdnro = getLastIdentity(objConn, "gti_achdiario")
    
    For i = 1 To CantidadTE
        If Usa_TE(i) Then
            StrSql = "INSERT INTO gti_achdiario_estr "
            StrSql = StrSql & "(achdnro, tenro, estrnro, achdfecha) "
            StrSql = StrSql & " VALUES ("
            StrSql = StrSql & l_achdnro & ","
            StrSql = StrSql & Estructura(i).Tenro & ","
            StrSql = StrSql & Estructura(i).Estrnro_Original & ","
            StrSql = StrSql & ConvFecha(p_fecha) & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    Next i
End If


FIN:
'Cierro todo y libero
If rs_Partes.State = adStateOpen Then rs_Partes.Close
If rs.State = adStateOpen Then rs.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_Partes_Estr.State = adStateOpen Then rs_Partes_Estr.Close

Set rs = Nothing
Set rs_Partes = Nothing
Set rs_Partes_Estr = Nothing
Set rs_Estructura = Nothing
End Sub


Public Sub Desglose_Jornada_Ausentismo(NroTer As Long, p_fecha As Date)

'Dim StrSql As String
Dim reghorario As Integer
Dim categoria As Integer
Dim producto As Integer
Dim PropJornalProduccion As Single
Dim l_achdnro As Long
Dim auxi As String
Dim rs As New ADODB.Recordset
Dim HorasProduccion As Integer
Dim HorasAusDia As Integer
Dim HorasTrabajadas As Single
Dim HorasTurno As Single

Dim horres As Single
Dim TotHorHHMM As String


'Borro para reprocesar
'StrSql = "delete from gti_achdiario_estr where achdnro in(" & _
'"SELECT achdnro FROM gti_achdiario WHERE ternro = " & NroTer & _
'" AND achdfecha = " & ConvFecha(p_fecha) & _
'" AND achdmanual = 0)"
'objConn.Execute StrSql, , adExecuteNoRecords
'
'StrSql = "delete from gti_achdiario where ternro = " & NroTer & _
'" AND achdfecha = " & ConvFecha(p_fecha) & " AND achdmanual = 0"
'objConn.Execute StrSql, , adExecuteNoRecords

'Busco las horas Produccion
StrSql = "SELECT confval FROM confrep WHERE repnro = 53 AND" & _
" conftipo = 'TH' AND confnrocol = 7"
OpenRecordset StrSql, rs
If rs.EOF Then
    Exit Sub
End If
HorasProduccion = rs!confval

'Busco el total de horas producción del día
StrSql = "SELECT * FROM gti_acumdiario " & _
    " WHERE (ternro = " & NroTer & ") AND " & _
    " (adfecha = " & ConvFecha(p_fecha) & ")" & _
    " AND thnro = " & HorasProduccion
OpenRecordset StrSql, rs
If rs.EOF Then
    Exit Sub
End If

HorasTrabajadas = rs!adcanthoras

'Busco las horas a desglosar en el confrep
StrSql = "SELECT confval FROM confrep WHERE repnro = 53 AND" & _
" conftipo = 'TH' AND confnrocol = 6"
OpenRecordset StrSql, rs
If rs.EOF Then
  Exit Sub
End If
HorasAusDia = rs!confval

' Busco en el histórico las estructuras
StrSql = " SELECT estrnro FROM his_estructura"
StrSql = StrSql & " WHERE his_estructura.tenro = 3 "
StrSql = StrSql & " AND his_estructura.ternro = " & NroTer
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(p_fecha)
StrSql = StrSql & " AND ((" & ConvFecha(p_fecha) & " <= htethasta) OR (htethasta is null))"
OpenRecordset StrSql, rs
If rs.EOF Then
  Exit Sub
End If
categoria = rs!estrnro

StrSql = " SELECT estrnro FROM his_estructura"
StrSql = StrSql & " WHERE his_estructura.tenro = 21 "
StrSql = StrSql & " AND his_estructura.ternro = " & NroTer
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(p_fecha)
StrSql = StrSql & " AND ((" & ConvFecha(p_fecha) & " <= htethasta) OR (htethasta is null))"
OpenRecordset StrSql, rs
If rs.EOF Then
  Exit Sub
End If
reghorario = rs!estrnro

StrSql = " SELECT estrnro FROM his_estructura"
StrSql = StrSql & " WHERE his_estructura.tenro = 38 "
StrSql = StrSql & " AND his_estructura.ternro = " & NroTer
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(p_fecha)
StrSql = StrSql & " AND ((" & ConvFecha(p_fecha) & " <= htethasta) OR (htethasta is null))"
OpenRecordset StrSql, rs
If rs.EOF Then
   Exit Sub
End If
producto = rs!estrnro

'Busco las horas producción del día
StrSql = "SELECT * FROM gti_acumdiario " & _
    " WHERE (ternro = " & NroTer & ") AND " & _
    " (adfecha = " & ConvFecha(p_fecha) & ")" & _
    " AND thnro = " & HorasAusDia
OpenRecordset StrSql, rs
   
If Not rs.EOF Then
    
    If (rs!adcanthoras = 0) Or (HorasTrabajadas = 0) Then Exit Sub
        
    HorasTurno = HorasTrabajadas / rs!adcanthoras
    PropJornalProduccion = Round(HorasTrabajadas / HorasTurno, 2)
    
    'Busco los partes de movilidad del día
    StrSql = " SELECT DISTINCT ternro, gmdnro, gmdhoras FROM gti_movdet "
    StrSql = StrSql & " WHERE gmdfecdesde <= " & ConvFecha(p_fecha)
    StrSql = StrSql & " AND " & ConvFecha(p_fecha) & " <= gmdfechasta "
    StrSql = StrSql & " AND ternro = " & NroTer
    OpenRecordset StrSql, objRs
    
    ' Si existen entonces
    Do While Not objRs.EOF
        'Hago el desgloce en las estructuras especificadas en el parte
        'de movilidad, por el tiempo definido en el.
    
        PropJornalProduccion = PropJornalProduccion - Round(objRs!gmdhoras / HorasTurno, 2)
        
        TotHorHHMM = CHoras(Round(objRs!gmdhoras / HorasTurno, 2), 60)
        ' Inserto en la tabla de desglose con horas producción de jornada
        StrSql = "INSERT INTO gti_achdiario "
        StrSql = StrSql & "(horas,achdcanthoras, achdfecha, achdmanual, achdvalido,ternro,thnro) "
        StrSql = StrSql & " VALUES (" & TotHorHHMM & "," & Round(objRs!gmdhoras / HorasTurno, 2) & ","
        StrSql = StrSql & ConvFecha(p_fecha) & ",0,-1,"
        StrSql = StrSql & NroTer & ", " & HorasAusDia & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'l_achdnro = getLastIdentity(objConn, "GTI_ACHDIARIO")
                
        StrSql = "select max(achdnro) as next_id from GTI_ACHDIARIO"
        OpenRecordset StrSql, rs
        l_achdnro = rs!next_id
        
        'Dentro del parte, busco una referencia a la estructura empaque
        StrSql = "SELECT * FROM gti_movdet_estr WHERE gmdnro = " & objRs!gmdnro
        StrSql = StrSql & " AND tenro = 21"
        OpenRecordset StrSql, rs

        StrSql = "INSERT INTO gti_achdiario_estr "
        StrSql = StrSql & "(achdnro, tenro, estrnro, achdfecha) "
        StrSql = StrSql & " VALUES (" & l_achdnro & ",21"
        'Reemplazo el empaque del empleado por el del parte
        If Not rs.EOF Then
            StrSql = StrSql & ", " & rs!estrnro & ","
        Else
            StrSql = StrSql & ", " & reghorario & ","
        End If
        StrSql = StrSql & ConvFecha(p_fecha) & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'Dentro del parte, busco una referencia a la estructura producto
        StrSql = "SELECT * FROM gti_movdet_estr WHERE gmdnro = " & objRs!gmdnro
        StrSql = StrSql & " AND tenro = 38"
        OpenRecordset StrSql, rs
        
        StrSql = "INSERT INTO gti_achdiario_estr "
        StrSql = StrSql & "(achdnro, tenro, estrnro, achdfecha)"
        StrSql = StrSql & " VALUES (" & l_achdnro & ",38"
        'Reemplazo el producto del empleado por el del parte
        If Not rs.EOF Then
            StrSql = StrSql & "," & rs!estrnro & ","
        Else
            StrSql = StrSql & "," & producto & ","
        End If
        StrSql = StrSql & ConvFecha(p_fecha) & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'Dentro del parte, busco una referencia a la estructura categoria
        StrSql = "SELECT * FROM gti_movdet_estr WHERE gmdnro = " & objRs!gmdnro
        StrSql = StrSql & " AND tenro = 3"
        OpenRecordset StrSql, rs
        
        StrSql = "INSERT INTO gti_achdiario_estr "
        StrSql = StrSql & "(achdnro, tenro, estrnro, achdfecha) "
        StrSql = StrSql & " VALUES (" & l_achdnro & ",3"
        'Reemplazo la categoria del empleado por el del parte
        If Not rs.EOF Then
            StrSql = StrSql & ", " & rs!estrnro & ","
        Else
            StrSql = StrSql & ", " & categoria & ","
        End If
        StrSql = StrSql & ConvFecha(p_fecha) & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        objRs.MoveNext
    Loop
    
    ' si aún quedan horas por desglosar, entonces uso las estructuras
    ' por descarte
    If PropJornalProduccion > 0 Then
        TotHorHHMM = CHoras(PropJornalProduccion, 60)
        
        ' Inserto en la tabla de desglose
        StrSql = "INSERT INTO gti_achdiario "
        StrSql = StrSql & "(horas, achdcanthoras, achdfecha, achdmanual, achdvalido,ternro,thnro) "
        StrSql = StrSql & " VALUES (" & TotHorHHMM & "," & PropJornalProduccion & ","
        StrSql = StrSql & ConvFecha(p_fecha) & ",0,-1,"
        StrSql = StrSql & NroTer & "," & HorasAusDia & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
               
        l_achdnro = getLastIdentity(objConn, "GTI_ACHDIARIO")
                
        StrSql = "select max(achdnro) as next_id from GTI_ACHDIARIO"
        OpenRecordset StrSql, rs
        l_achdnro = rs!next_id
        
        
        'Inserto en la tabla de desgloce por estructura
        '(un registro por cada estructura)
        
        StrSql = "INSERT INTO gti_achdiario_estr "
        StrSql = StrSql & "(achdnro, tenro, estrnro, achdfecha) "
        StrSql = StrSql & " VALUES (" & l_achdnro & ",3"
        StrSql = StrSql & ", " & categoria & ","
        StrSql = StrSql & ConvFecha(p_fecha) & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
   
        StrSql = "INSERT INTO gti_achdiario_estr "
        StrSql = StrSql & "(achdnro, tenro, estrnro, achdfecha) "
        StrSql = StrSql & " VALUES (" & l_achdnro & ",21"
        StrSql = StrSql & ", " & reghorario & ","
        StrSql = StrSql & ConvFecha(p_fecha) & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    
        StrSql = "INSERT INTO gti_achdiario_estr "
        StrSql = StrSql & "(achdnro, tenro, estrnro, achdfecha) "
        StrSql = StrSql & " VALUES (" & l_achdnro & ",38"
        StrSql = StrSql & ", " & producto & ","
        StrSql = StrSql & ConvFecha(p_fecha) & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        rs.Close
    End If
    
End If

If objRs.State = adStateOpen Then objRs.Close
If rs.State = adStateOpen Then rs.Close

Set rs = Nothing

End Sub




Public Sub Desglose_Jornada_Ausentismo_Nuevo(ByVal NroTer As Long, ByVal p_fecha As Date)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de genera el desglose de Jornada de Ausentismo.
' Autor      : FGZ
' Fecha      : 17/11/2005
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim Estructura(1 To 5) As TDesglose
Dim Usa_TE(1 To 5) As Boolean
Dim CantidadTE As Integer
Dim i As Integer

Dim THorasProduccion As Long
Dim THorasProdDia As Long
Dim HorasTrabajadas As Single
Dim HorasTrabajadasDia As Single

Dim PropJornalProduccion As Single
Dim l_achdnro As Long
Dim HorasTurno As Single

Dim rs As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Partes As New ADODB.Recordset
Dim rs_Partes_Estr As New ADODB.Recordset
Dim Seguir As Boolean
Dim TotHorHHMM As String
Dim rs_FT As New ADODB.Recordset


''Borro para reprocesar
'StrSql = "DELETE FROM gti_achdiario_estr "
'StrSql = StrSql & " WHERE achdnro IN( "
'    StrSql = StrSql & " SELECT achdnro FROM gti_achdiario "
'    StrSql = StrSql & " WHERE (ternro = " & NroTer
'    StrSql = StrSql & " AND achdfecha = " & ConvFecha(p_fecha)
'    StrSql = StrSql & " AND achdmanual = 0) "
'StrSql = StrSql & ")"
'objConn.Execute StrSql, , adExecuteNoRecords
'
'StrSql = "DELETE FROM gti_achdiario "
'StrSql = StrSql & " WHERE ternro = " & NroTer
'StrSql = StrSql & " AND achdfecha = " & ConvFecha(p_fecha)
'StrSql = StrSql & " AND achdmanual = 0"
'objConn.Execute StrSql, , adExecuteNoRecords


If depurar Then
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Desglose_Jornada_Ausentismo"
End If

'Busco las horas Produccion
StrSql = "SELECT confval FROM confrep "
StrSql = StrSql & " WHERE repnro = 53 "
StrSql = StrSql & " AND conftipo = 'TH' AND confnrocol = 7"
If rs.State = adStateOpen Then rs.Close
OpenRecordset StrSql, rs
If rs.EOF Then
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No esta configurado el tipo de hora produccion. Confrep 53 columna 4."
    End If
    Exit Sub
End If
THorasProduccion = rs!confval

'Busco el total de horas producción del día
StrSql = "SELECT adcanthoras FROM gti_acumdiario "
StrSql = StrSql & " WHERE ternro = " & NroTer
StrSql = StrSql & " AND adfecha = " & ConvFecha(p_fecha)
StrSql = StrSql & " AND thnro = " & THorasProduccion
If rs.State = adStateOpen Then rs.Close
OpenRecordset StrSql, rs
If rs.EOF Then
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No hay tipo de hora " & THorasProduccion & " para la fecha " & p_fecha
    End If
    Exit Sub
End If
HorasTrabajadas = rs!adcanthoras

'Busco las horas a desglosar en el confrep
StrSql = "SELECT confval FROM confrep WHERE repnro = 53 "
StrSql = StrSql & " AND conftipo = 'TH' AND confnrocol = 6"
If rs.State = adStateOpen Then rs.Close
OpenRecordset StrSql, rs
If rs.EOF Then
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No esta configurado el tipo de hora a desglozar. Confrep 53 columna 5."
    End If
    Exit Sub
End If
THorasProdDia = rs!confval

'Busco los tipos de estructuras a Desglozar
StrSql = "SELECT confval, confnrocol FROM confrep WHERE repnro = 53 "
StrSql = StrSql & " AND confnrocol >= 50 AND confnrocol <= 54"
StrSql = StrSql & " ORDER BY confnrocol"
If rs.State = adStateOpen Then rs.Close
OpenRecordset StrSql, rs
If rs.EOF Then
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No hay configurado ningun tipo de estructura a desglozar."
    End If
    Exit Sub
End If

'Inicializo las estructuras del desglose
For i = 1 To 5
    Estructura(i).Tenro = 0
    Estructura(i).Estrnro_Original = 0
    Usa_TE(i) = False
Next i

i = 0
Do While Not rs.EOF And i <= 5
    If Not EsNulo(rs!confval) Then
        i = i + 1
        Estructura(i).Tenro = rs!confval
        Usa_TE(i) = True
    
        'Busco en el histórico las estructuras
        StrSql = " SELECT estrnro FROM his_estructura"
        StrSql = StrSql & " WHERE his_estructura.tenro = " & Estructura(i).Tenro
        StrSql = StrSql & " AND his_estructura.ternro = " & NroTer
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(p_fecha)
        StrSql = StrSql & " AND ((" & ConvFecha(p_fecha) & " <= htethasta) OR (htethasta is null))"
        If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
        OpenRecordset StrSql, rs_Estructura
        If rs_Estructura.EOF Then
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "El empleado no tiene activa ningun estructura de tipo " & Estructura(i).Tenro & "."
            End If
            'Exit Sub
        Else
            Estructura(i).Estrnro_Original = rs_Estructura!estrnro
        End If
    End If

    rs.MoveNext
Loop
CantidadTE = i


'FGZ - 10/07/2008 - Le agregué este control porque sino insertaba registros con estrnro = 0
Seguir = True
For i = 1 To 5
    If Estructura(i).Tenro <> 0 Then
        If Estructura(i).Estrnro_Original = 0 Then
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "El empleado no tiene activa ningun estructura de tipo " & Estructura(i).Tenro & "."
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No se puede hacer ningun desgloce."
            End If
            Seguir = False
        End If
    End If
Next i
If Not Seguir Then
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "El empleado no tiene activa alguna estructura necesaria para el desgloce."
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No se efectuará ningun desgloce."
    End If
    GoTo FIN
End If

'Busco las horas producción del día
StrSql = "SELECT adcanthoras FROM gti_acumdiario "
StrSql = StrSql & " WHERE ternro = " & NroTer
StrSql = StrSql & " AND adfecha = " & ConvFecha(p_fecha)
StrSql = StrSql & " AND thnro = " & THorasProdDia
If rs.State = adStateOpen Then rs.Close
OpenRecordset StrSql, rs
If rs.EOF Then
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "El empleado no tiene horas produccion (" & THorasProdDia & ") en la fecha " & p_fecha
    End If
    'esto es idea mia, poner cantidad = 1 por default
    'HorasTrabajadasDia = 1
    Exit Sub
Else
    HorasTrabajadasDia = rs!adcanthoras
End If
If (HorasTrabajadasDia = 0) Or (HorasTrabajadas = 0) Then
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "La cantidad de horas 0."
    End If
    Exit Sub
End If
    
HorasTurno = HorasTrabajadas / HorasTrabajadasDia
PropJornalProduccion = Round(HorasTrabajadas / HorasTurno, 2)

'Busco los partes de movilidad del día
'StrSql = " SELECT DISTINCT ternro, gmdnro, gmdhoras FROM gti_movdet "
'StrSql = StrSql & " WHERE gmdfecdesde <= " & ConvFecha(p_fecha)
'StrSql = StrSql & " AND " & ConvFecha(p_fecha) & " <= gmdfechasta "
'StrSql = StrSql & " AND ternro = " & NroTer
'If rs_Partes.State = adStateOpen Then rs_Partes.Close
'OpenRecordset StrSql, rs_Partes

StrSql = " SELECT DISTINCT gti_movdet.gcpnro,gti_movdet.ternro, gti_movdet.gmdnro, gti_movdet.gmdhoras, gti_cabparte.ft, gti_cabparte.ftap FROM gti_cabparte "
StrSql = StrSql & " INNER JOIN gti_movdet ON gti_cabparte.gcpnro = gti_movdet.gcpnro "
StrSql = StrSql & " WHERE gmdfecdesde <= " & ConvFecha(p_fecha)
StrSql = StrSql & " AND " & ConvFecha(p_fecha) & " <= gmdfechasta "
StrSql = StrSql & " AND ternro = " & NroTer
StrSql = StrSql & " AND (ft = 0 OR (ft = -1 AND ftap = -1))"
OpenRecordset StrSql, rs_Partes
Do While Not rs_Partes.EOF
    'Hago el desgloce en las estructuras especificadas en el parte
    'de movilidad, por el tiempo definido en el.

    StrSql = "SELECT input_ft.idnro,input_ft.origen, gti_cabparte.ft, gti_cabparte.ftap FROM input_ft "
    StrSql = StrSql & " INNER JOIN gti_cabparte ON input_ft.origen = gti_cabparte.gcpnro "
    StrSql = StrSql & " WHERE idtipoinput = 9 "
    StrSql = StrSql & " AND origen = " & rs_Partes!GCPNRO
    OpenRecordset StrSql, rs_FT
    If Not rs_FT.EOF Then
        If rs_FT!ftap = -1 Then
            Call InsertarFT(rs_FT!idnro, 9, rs_FT!Origen)
        End If
    End If

    PropJornalProduccion = PropJornalProduccion - Round(rs_Partes!gmdhoras / HorasTurno, 2)
        
    TotHorHHMM = CHoras(Round(rs_Partes!gmdhoras / HorasTurno, 2), 60)
    'Inserto en la tabla de desglose con horas producción de jornada
    StrSql = "INSERT INTO gti_achdiario "
    StrSql = StrSql & "(horas, achdcanthoras, achdfecha, achdmanual, achdvalido,ternro,thnro) "
    StrSql = StrSql & " VALUES (" & TotHorHHMM & "," & Round(rs_Partes!gmdhoras / HorasTurno, 2) & ","
    StrSql = StrSql & ConvFecha(p_fecha) & ",0,-1,"
    StrSql = StrSql & NroTer & ", " & THorasProdDia & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    l_achdnro = getLastIdentity(objConn, "gti_achdiario")
            
    'busco dentro del parte cada TE a desglozar
    For i = 1 To CantidadTE
        If Usa_TE(i) Then
            StrSql = "SELECT * FROM gti_movdet_estr WHERE gmdnro = " & rs_Partes!gmdnro
            StrSql = StrSql & " AND tenro = " & Estructura(i).Tenro
            If rs_Partes_Estr.State = adStateOpen Then rs_Partes_Estr.Close
            OpenRecordset StrSql, rs_Partes_Estr
            
            StrSql = "INSERT INTO gti_achdiario_estr "
            StrSql = StrSql & "(achdnro, tenro, estrnro, achdfecha) "
            StrSql = StrSql & " VALUES ("
            StrSql = StrSql & l_achdnro & ","
            StrSql = StrSql & Estructura(i).Tenro & ","
            If Not rs_Partes_Estr.EOF Then
                StrSql = StrSql & rs_Partes_Estr!estrnro & ","
            Else
                'Reemplazo el empaque del empleado por el del parte
                StrSql = StrSql & Estructura(i).Estrnro_Original & ","
            End If
            StrSql = StrSql & ConvFecha(p_fecha) & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    Next i
    
    rs_Partes.MoveNext
Loop


' si aún quedan horas por desglosar, entonces uso las estructuras
' por descarte
If PropJornalProduccion > 0 Then
    ' Inserto en la tabla de desglose
    
    TotHorHHMM = CHoras(PropJornalProduccion, 60)
    StrSql = "INSERT INTO gti_achdiario "
    StrSql = StrSql & "(horas, achdcanthoras, achdfecha, achdmanual, achdvalido,ternro,thnro) "
    StrSql = StrSql & " VALUES (" & TotHorHHMM & "," & PropJornalProduccion & ","
    StrSql = StrSql & ConvFecha(p_fecha) & ",0,-1,"
    StrSql = StrSql & NroTer & "," & THorasProdDia & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    l_achdnro = getLastIdentity(objConn, "gti_achdiario")
    
    For i = 1 To CantidadTE
        If Usa_TE(i) Then
            StrSql = "INSERT INTO gti_achdiario_estr "
            StrSql = StrSql & "(achdnro, tenro, estrnro, achdfecha) "
            StrSql = StrSql & " VALUES ("
            StrSql = StrSql & l_achdnro & ","
            StrSql = StrSql & Estructura(i).Tenro & ","
            StrSql = StrSql & Estructura(i).Estrnro_Original & ","
            StrSql = StrSql & ConvFecha(p_fecha) & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    Next i
End If

FIN:
'Cierro todo y libero
If rs_Partes.State = adStateOpen Then rs_Partes.Close
If rs.State = adStateOpen Then rs.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_Partes_Estr.State = adStateOpen Then rs_Partes_Estr.Close

Set rs = Nothing
Set rs_Partes = Nothing
Set rs_Partes_Estr = Nothing
Set rs_Estructura = Nothing
End Sub


Public Function IndiceTE(ByVal TE As Long, ByVal Estr_Ind) As Long
Dim J As Long
Dim Encontro As Boolean

J = 1
Encontro = False
While J <= 5 And Not Encontro
    If Estr_Ind(J) = TE Then
        Encontro = True
    Else
        J = J + 1
    End If
Wend
If Not Encontro Then
    IndiceTE = 0
Else
    IndiceTE = J
End If
End Function




