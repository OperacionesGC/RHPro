Attribute VB_Name = "mdlVersiones"

Option Explicit


'Version: 1.01  'Inicial

'Const Version = "1.01"    'Version Inicial de Vencimiento de dias Vacaciones.
'Const FechaVersion = "23/10/2009"
''           Politica 1512 - Vencimiento de vacaciones.


'Const Version = "1.02"
'Const FechaVersion = "16/11/2009" 'FGZ
''           Problema con la funcion de validacion de version.

'------------------------------------------------------------------------------------
'Const Version = "3.00"
'Const FechaVersion = "14/04/2010" 'FGZ
'           Ahora los periodos de vacaciones ahora pueden tener alcance por estructura
'           hubo que agregar 2 parametros mas

'Const Version = "3.01"
'Const FechaVersion = "04/10/2010" 'Margiotta, Emanuel
''           Se agrego la busqueda del periodo siguiente para pasar por parametro en la funcion DiasVencidos

'Const Version = "3.02"
'Const FechaVersion = "14/04/2011" 'Lisandro Moro
''           Se agrego el vencimiento version colombia

Global Const Version = "3.03"
Global Const FechaVersion = "03/09/2012" 'Margiotta Emanuel - CAS-13764 – Nivelación GIV
'          Se creo la versión para estandarizar el nro de Versión.

' ==================================================================================================================

Public Function ValidarVBD(ByVal Version As String, ByVal TipoProceso As Long, ByVal TipoBD As Integer, ByVal codPais As Integer) As Boolean

End Function
Public Function ValidarVersion(ByVal Version As String, ByVal TipoProceso As Long, ByVal TipoBD As Integer) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Validacion de estructura de BD
' Autor      : FGZ
' Fecha      : 06/08/2012
' ---------------------------------------------------------------------------------------------
Dim V As Boolean
Dim Texto As String
Dim rs As New ADODB.Recordset

On Error GoTo ME_Version

V = True

    If Version >= "1.01" Then
        'Revisar los campos
        'vacdiascor.venc
        Texto = "Revisar los campos: vacdiascor.venc"
        
        StrSql = "Select venc from vacdiascor where ternro = 1"
        OpenRecordset StrSql, rs

        V = True
    End If

    ValidarVersion = V
Exit Function

ME_Version:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Estructura de BD incompatible con la version del proceso."
    Flog.writeline Espacios(Tabulador * 1) & Texto
    Flog.writeline
    V = False
End Function

