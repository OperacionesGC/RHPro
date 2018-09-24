Attribute VB_Name = "MdlMidioma"
Option Explicit
' ---------------------------------------------------------------------------------------------
' Descripcion: Módulo que se encarga de las validaciones de Multi-Idioma.
' Autor      : Gonzalez Nicolás
' Fecha      : 19/01/2012
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Global MI As Boolean 'MultiIdioma -
Global Lenguaje As String 'MultiIdioma - Ej: esAR


Public Sub Valida_MultiIdiomaActivo(ByVal usuario)
' ---------------------------------------------------------------------------------------------
' Descripcion: Valida que este activo Multi-Idioma y Guarda el idioma predeterminado del usuario.
' Autor      : Gonzalez Nicolás
' Fecha      : 19/01/2012
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
 
 Dim rs_mi As New ADODB.Recordset
 Dim rs_mi2 As New ADODB.Recordset
 Dim val_siguiente As Boolean
 
 MI = False
 val_siguiente = False
 Lenguaje = ""
 
 'VALIDO QUE ESTE ACTIVA LA TRADUCCION
 StrSql = "SELECT confactivo FROM confper WHERE confnro = 8 AND confactivo = -1"
 OpenRecordset StrSql, rs_mi
 
 If rs_mi.EOF Then
     Flog.writeline "Multi-idioma se encuentra deshabilitado - Revisar Configuración de Empresa para activarlo - Código 8 "
 Else
    '____________________________
    'VALIDA EL TIPO DE BASE
    '----------------------------
    Select Case TipoBD
        Case 1 'db2
            StrSql = ""
        Case 2 ' Informix
            StrSql = ""
        Case 3 ' sql server
            '____________________________________
            'VALIDA QUE EXISTA LA TABLA lenguaje
            '------------------------------------
            StrSql = "SELECT TABLE_NAME FROM information_schema.tables WHERE table_name = 'lenguaje'"
            OpenRecordset StrSql, rs_mi2
            If Not rs_mi2.EOF Then
                rs_mi2.Close
                '____________________________________________
                'VALIDA QUE EXISTA LA TABLA lenguaje_etiqueta
                '--------------------------------------------
                StrSql = "SELECT TABLE_NAME FROM information_schema.tables WHERE table_name = 'lenguaje_etiqueta'"
                OpenRecordset StrSql, rs_mi2
                If Not rs_mi2.EOF Then
                    val_siguiente = True
                End If
                rs_mi2.Close
            End If
            '_______________________________________________________________________________________
            'SI ENCONTRO LAS 2 TABLAS ANTERIORES BUSCA EL IDIOMA DEL USUARIO QUE EJECUTO EL PROCESO.
            '---------------------------------------------------------------------------------------
            If val_siguiente = True Then
                '___________________________
                'BUSCO EL IDIOMA DEL USUARIO
                '---------------------------
                StrSql = "SELECT lencod FROM user_per "
                StrSql = StrSql & " INNER JOIN lenguaje ON lenguaje.lennro = user_per.lennro "
                StrSql = StrSql & " WHERE UPPER(iduser) = '" & UCase(usuario) & "'"
                OpenRecordset StrSql, rs_mi2
                If Not rs_mi2.EOF Then
                    Lenguaje = Trim(rs_mi2!lencod)
                    MI = True
                End If
                rs_mi2.Close
            End If
            
        Case 4 'Oracle
            '____________________________________
            'VALIDA QUE EXISTA LA TABLA lenguaje
            '------------------------------------
            StrSql = "select table_name from user_tables where lower(table_name) = 'lenguaje'"
            OpenRecordset StrSql, rs_mi2
            If Not rs_mi2.EOF Then
                rs_mi2.Close
                '____________________________________________
                'VALIDA QUE EXISTA LA TABLA lenguaje_etiqueta
                '--------------------------------------------
                StrSql = "select table_name from user_tables where lower(table_name) = 'lenguaje_etiqueta'"
                'StrSql = "SELECT TABLE_NAME FROM information_schema.tables WHERE table_name = 'lenguaje_etiqueta'"
                OpenRecordset StrSql, rs_mi2
                If Not rs_mi2.EOF Then
                    val_siguiente = True
                End If
                rs_mi2.Close
            End If
            '_______________________________________________________________________________________
            'SI ENCONTRO LAS 2 TABLAS ANTERIORES BUSCA EL IDIOMA DEL USUARIO QUE EJECUTO EL PROCESO.
            '---------------------------------------------------------------------------------------
            If val_siguiente = True Then
                '___________________________
                'BUSCO EL IDIOMA DEL USUARIO
                '---------------------------
                StrSql = "SELECT lencod FROM user_per "
                StrSql = StrSql & " INNER JOIN lenguaje ON lenguaje.lennro = user_per.lennro "
                StrSql = StrSql & " WHERE UPPER(iduser) = '" & UCase(usuario) & "'"
                OpenRecordset StrSql, rs_mi2
                If Not rs_mi2.EOF Then
                    Lenguaje = Trim(rs_mi2!lencod)
                End If
                rs_mi2.Close
            End If
        
    End Select
    
 End If
 
 rs_mi.Close
 

End Sub
Public Function EscribeLogMI(ByVal Texto)
' ---------------------------------------------------------------------------------------------
' Descripcion: Traduce el texto al idioma del usuario | Si no la encuentra devuelve texto original.
' Autor      : Gonzalez Nicolás
' Fecha      : 19/01/2012
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    Dim rs_mi As New ADODB.Recordset
    'EscribeLogMI = Texto
    EscribeLogMI = "{" & Texto & "}"
    
    If MI = True And Lenguaje <> "" Then
        StrSql = "SELECT " & Lenguaje & " FROM lenguaje_etiqueta "
        StrSql = StrSql & " WHERE etiqueta = '" & Texto & "'"
        OpenRecordset StrSql, rs_mi
        If Not rs_mi.EOF Then
            If Not EsNulo(rs_mi(0)) Then
                'EscribeLogMI = rs_mi(0)
                EscribeLogMI = "[" & rs_mi(0) & "]"
            End If
        End If
        rs_mi.Close
    End If

End Function

