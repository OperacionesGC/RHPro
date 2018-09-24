Attribute VB_Name = "MdlBatchLog"
Option Explicit
' ---------------------------------------------------------------------------------------------
' Descripcion: Modulo BatchLog. Procedimientos y Funciones para grabar y borrar en la tabla batch_logs
' Autor      : Martin Ferraro
' Fecha      : 07/06/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------


Public Sub bl_insertar(ByVal bpronro As Long, ByVal tipo As Integer, ByVal desabr As String, ByVal desext As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta en la tabla batch_logs
' Autor      : Martin Ferraro
' Fecha      : 07/06/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

    StrSql = " INSERT INTO batch_log(bpronro,tipo,desabr,desext) "
    StrSql = StrSql & " VALUES ( "
    StrSql = StrSql & "  " & bpronro
    StrSql = StrSql & " ," & tipo
    StrSql = StrSql & " ,'" & Mid(desabr, 1, 100) & "' "
    StrSql = StrSql & " ,'" & Mid(desext, 1, 1000) & "' "
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub


Public Sub bl_insertar_con_ternro(ByVal ternro As Long, ByVal bpronro As Long, ByVal tipo As Integer, ByVal desabr As String, ByVal desext As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta en la tabla batch_logs concatenando la descr del empleado en desabr
' Autor      : Martin Ferraro
' Fecha      : 07/06/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_reg As New ADODB.Recordset
Dim Empleado As String

    'Busco al empleado
    Empleado = ""
    StrSql = "SELECT empleg, terape, ternom FROM v_empleado WHERE v_empleado.ternro = " & ternro
    OpenRecordset StrSql, rs_reg
    If Not rs_reg.EOF Then
        Empleado = rs_reg!empleg & " - " & rs_reg!terape & ", " & rs_reg!ternom & ": "
    End If
    rs_reg.Close
    
    desabr = Empleado & desabr

    StrSql = " INSERT INTO batch_log(bpronro,tipo,desabr,desext) "
    StrSql = StrSql & " VALUES ( "
    StrSql = StrSql & "  " & bpronro
    StrSql = StrSql & " ," & tipo
    StrSql = StrSql & " ,'" & Mid(desabr, 1, 100) & "' "
    StrSql = StrSql & " ,'" & Mid(desext, 1, 1000) & "' "
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Set rs_reg = Nothing

End Sub


Public Sub bl_borrar(ByVal bpronro As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Borra todos los registro con el bpronro en la tabla batch_logs
' Autor      : Martin Ferraro
' Fecha      : 07/06/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

    StrSql = " DELETE FROM batch_log WHERE bpronro = " & bpronro
    objConn.Execute StrSql, , adExecuteNoRecords
    
End Sub
