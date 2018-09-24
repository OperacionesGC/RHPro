Attribute VB_Name = "MdlPoliticasLiq"
' Modulo de Politicas de Liquidacion

Option Explicit
Global PoliticaOK As Boolean 'si cargo bien o no la politica llamada


' Debe haber 1 Variable por cada tipo de parametro posible
Global st_Opcion As Integer
Global st_VentSal As String
Global st_VentEnt As String
Global st_Iteraciones As Integer
Global st_Tolerancia As String
Global st_TipoHora1 As Long
Global st_Distancia As Integer
Global st_TamañoVentana As String
Global st_TipoDia1 As Integer
Global st_CantidadDias As Integer
Global st_FactorDivision As Integer
Global st_Escala As Integer
Global st_ModeloPago As Integer
Global st_ModeloDto As Integer

Public Sub SetearParametrosPolitica(ByVal Detalle As Long)
Dim rsPolitica As New ADODB.Recordset

    PoliticaOK = False
    StrSql = " SELECT * FROM gti_pol_det_param " & _
             " INNER JOIN gti_pol_param ON gti_pol_det_param.polparamnro = gti_pol_param.polparamnro " & _
             " WHERE detpolnro = " & Detalle & _
             " ORDER BY gti_pol_param.polparamnro"
    OpenRecordset StrSql, rsPolitica

    If Not rsPolitica.EOF Then
        PoliticaOK = True
    End If

    Do While Not rsPolitica.EOF
        Select Case rsPolitica!polparamnro
        Case 1:
            st_Opcion = CInt(rsPolitica!polparamvalor)
        Case 2:
            st_VentSal = Format(rsPolitica!polparamvalor, "0000")
        Case 3:
            ' por ahora esta vacio
        Case 4:
            st_VentEnt = Format(rsPolitica!polparamvalor, "0000")
        Case 5:
            st_Iteraciones = CInt(rsPolitica!polparamvalor)
        Case 6:
            st_Tolerancia = Format(rsPolitica!polparamvalor, "0000")
        Case 7:
            st_Distancia = CInt(rsPolitica!polparamvalor)
        Case 8:
            st_TipoHora1 = CLng(rsPolitica!polparamvalor)
        Case 9:
            st_TamañoVentana = Format(rsPolitica!polparamvalor, "0000")
        Case 10:
            st_TipoDia1 = CInt(rsPolitica!polparamvalor)
        Case 11:
            st_CantidadDias = CInt(rsPolitica!polparamvalor)
        Case 12:
            st_FactorDivision = CInt(rsPolitica!polparamvalor)
        Case 13:
            st_Escala = CInt(rsPolitica!polparamvalor)
        Case 14:
            st_ModeloPago = CInt(rsPolitica!polparamvalor)
        Case 15:
            st_ModeloDto = CInt(rsPolitica!polparamvalor)
        Case Else
        
        End Select
        
        rsPolitica.MoveNext
    Loop
    
End Sub



Public Sub Politica(ByVal Numero As Long, ByVal P_Fecha As Date, ByRef Politica_OK As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca el nro de grilla configurado en la politica 1502 de Vacaciones segun su alcance.
' Autor      : FGZ
' Fecha      : 28/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Politica As New ADODB.Recordset

Dim det As Integer
Dim Cabecera As Long
Dim Detalle As Long

    PoliticaOK = False
    
    StrSql = "SELECT gti_cabpolitica.cabpolnro,gti_cabpolitica.cabpolnivel, gti_alcanpolitica.alcpolnivel, gti_alcanpolitica.alcpolorigen, gti_detpolitica.detpolnro, gti_detpolitica.detpolprograma " & _
        "FROM gti_cabpolitica INNER JOIN gti_alcanpolitica ON gti_cabpolitica.cabpolnro = gti_alcanpolitica.cabpolnro INNER JOIN gti_detpolitica ON gti_alcanpolitica.detpolnro = gti_detpolitica.detpolnro " & _
        "WHERE gti_cabpolitica.cabpolnivel = " & Numero & " And gti_alcanpolitica.alcpolnivel = 3 And gti_alcanpolitica.alcpolorigen = " & buliq_empleado!ternro & " AND gti_cabpolitica.cabpolestado = -1 And gti_alcanpolitica.alcpolestado = -1 "
    OpenRecordset StrSql, rs_Politica
    
    If rs_Politica.EOF Then
        StrSql = " SELECT gti_cabpolitica.cabpolnro, gti_cabpolitica.cabpolnivel,gti_alcanpolitica.alcpolnivel, gti_alcanpolitica.alcpolorigen, gti_detpolitica.detpolnro,gti_detpolitica.detpolprograma,alcance_testr.alteOrden "
        StrSql = StrSql & " FROM gti_cabpolitica "
        StrSql = StrSql & " INNER JOIN gti_alcanpolitica ON gti_cabpolitica.cabpolnro = gti_alcanpolitica.cabpolnro "
        StrSql = StrSql & " INNER JOIN gti_detpolitica ON gti_alcanpolitica.detpolnro = gti_detpolitica.detpolnro "
        StrSql = StrSql & " INNER JOIN his_estructura ON gti_alcanpolitica.alcpolorigen = his_estructura.estrnro "
        StrSql = StrSql & " INNER JOIN alcance_tEstr ON his_estructura.tenro = alcance_tEstr.tenro "
        StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = his_estructura.ternro"
        StrSql = StrSql & " WHERE gti_cabpolitica.cabpolnivel = " & Numero
        StrSql = StrSql & " And gti_alcanpolitica.alcpolnivel = 2 "
        StrSql = StrSql & " And gti_cabpolitica.cabpolestado = -1 "
        StrSql = StrSql & " And gti_alcanpolitica.alcpolestado = -1 "
        StrSql = StrSql & " And alcance_testr.tanro = 1 "
        StrSql = StrSql & " And empleado.ternro = " & buliq_empleado!ternro
        StrSql = StrSql & " And (his_estructura.htetdesde <= " & ConvFecha(P_Fecha) & ")"
        StrSql = StrSql & " And ((" & ConvFecha(P_Fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
        StrSql = StrSql & " ORDER BY alcance_testr.AlteOrden Asc "
        If rs_Politica.State = adStateOpen Then rs_Politica.Close
        OpenRecordset StrSql, rs_Politica
        
        If rs_Politica.EOF Then
            StrSql = " SELECT gti_cabpolitica.cabpolnro,gti_cabpolitica.cabpolnivel, gti_alcanpolitica.alcpolnivel, gti_alcanpolitica.alcpolorigen, gti_detpolitica.detpolnro, gti_detpolitica.detpolprograma " & _
            " FROM gti_cabpolitica INNER JOIN gti_alcanpolitica ON gti_cabpolitica.cabpolnro = gti_alcanpolitica.cabpolnro INNER JOIN gti_detpolitica ON gti_alcanpolitica.detpolnro = gti_detpolitica.detpolnro " & _
            " WHERE gti_cabpolitica.cabpolnivel = " & Numero & " And gti_alcanpolitica.alcpolnivel = 1 And gti_cabpolitica.cabpolestado = -1 And gti_alcanpolitica.alcpolestado = -1 "
            If rs_Politica.State = adStateOpen Then rs_Politica.Close
            OpenRecordset StrSql, rs_Politica
        End If
    End If
    
    If Not rs_Politica.EOF Then
        det = rs_Politica!detpolprograma
        Cabecera = rs_Politica!cabpolnro
        Detalle = rs_Politica!detpolnro
        
        Select Case Numero
        Case 1501: 'Proporcion de dias de Vacaciones
            Call politica1501(det, Cabecera, Detalle)
        Case 1502: 'Escala
            Call politica1502(det, Cabecera, Detalle)
            If PoliticaOK Then
            Else
                PoliticaOK = False
            End If
                           
        End Select
    End If

    Politica_OK = PoliticaOK
    If rs_Politica.State = adStateOpen Then rs_Politica.Close
    Set rs_Politica = Nothing
    
End Sub


Private Sub politica1502(ByVal subn As Long, ByVal Cabecera As Long, ByVal Detalle As Long)
    Call SetearParametrosPolitica(Detalle)
End Sub


Private Sub politica1501(ByVal subn As Long, ByVal Cabecera As Long, ByVal Detalle As Long)
    Call SetearParametrosPolitica(Detalle)
End Sub
