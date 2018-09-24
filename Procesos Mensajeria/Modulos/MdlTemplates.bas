Attribute VB_Name = "MdlTemplates"
'-----------------------------------------------------------------------------
'Archivo        : MdlTemplates
'Creador        : Lisandro Moro
'Fecha Creacion : 10/09/2010
'Descripcion    : Genera los templates para los emails
'Modificacion   :
'----------------------------------------------------------------------------------------------------------
'Defino una expresion regular para definir el formato de las etiquetas.
Global m_tagrex As String
'm_tagrex = "\{ÑÑ[A-Z]{1,2}([0-9]{1,3}|[0-9]{1,3}\_[0-9]{1,2})\}"
'Busca lo que encuentre entre {} Y contenga
'Los valores que comiencen con ÑÑ
' Y a continuacion hayan 1 o 2 letras
' Y (a continuacion hayan 1 a 3 numeros) O (a continuacion hayan 1 a 3 numeros seguidos de un guion bajo seguido de 1 o 2 numeros )
' Ej: {ÑÑA001}, {ÑÑAB1}, {ÑÑAD001_10}, {ÑÑA100_6}

'Defino los separadores a utilizar al momento de generar el string con los valores
'Ej: "@@{ÑÑA001}::Nombre Empleado@@{ÑÑA002}::Apellido Empleado@@"
Global Const septptepar As String = "::"
Global Const septptegrp As String = "@@"
Global AnexoImgHeader As String
'l_septptepar = "::"
'l_septptegrp = "@@"
'response.Write ("<script>")
'response.Write ("var septptepar = '" & l_septptepar & "';")
'response.Write ("var septptegrp = '" & l_septptegrp & "';")
'response.Write ("</script>")


' ________________________________________________________________________________________
'10-06-2011 - Leticia A. - Se modifico la forma de buscar el directorio de Templates, ahora se usa la dirección de sis_direntradas
' ________________________________________________________________________________________
  'sub generarEmailTemplate(tipoOrigen, origen, funcion, ArrTags, archivo ,asunto ,msgbody ,destino ,CC ,CCO)
Function generarTemplate(tipoOrigen, Origen, funcion, ArrTags, dirSistema, codproc, ByVal Col1 As Long, fileName) As String
Dim Ok As Boolean
Dim rs_tiponoti As New ADODB.Recordset

    'ArrTags: Arreglo separado por "@@" el par y por "::" la etiqueta y el valor
    m_tagrex = "\{ÑÑ[A-Z]{1,2}([0-9]{1,3}|[0-9]{1,3}\_[0-9]{1,2})\}"
    'septptepar = "::"
    'septptegrp = "@@"

    generarTemplate = ""
    If (tipoOrigen = "" Or Origen = "" Or funcion = "" Or dirSistema = "" Or codproc = "") Then 'OR ArrTags = ""
        generarTemplate = ""
        Exit Function
    End If

    Dim m_sql As String
    Dim m_rs  As New ADODB.Recordset

    Dim m_archivo As String
    Dim m_dirSistema As String
    Dim m_dirTemplate As String
    Dim m_fileTemplate As String
    Dim m_contenido As String
    
        
    BuscarDirectorioT m_dirSistema, Ok
    
    If Ok = False Then
        generarTemplate = ""
        Exit Function
    End If
    
    
    'Busco la carpeta y el archivo del template
    m_sql = " SELECT tpltefile, tpltefolder "
    m_sql = m_sql & " FROM tpte_files "
    m_sql = m_sql & " WHERE ttorigennro = " & tipoOrigen
    m_sql = m_sql & " AND tplteorigen = " & Origen
    m_sql = m_sql & " AND tfunnro = " & funcion
    OpenRecordset m_sql, m_rs
    If Not m_rs.EOF Then
        m_dirTemplate = m_rs("tpltefolder")
        m_fileTemplate = m_rs("tpltefile")
        m_archivo = m_dirSistema & "\" & m_dirTemplate & "\" & m_fileTemplate
    Else
        m_archivo = ""
        generarTemplate = ""
        Exit Function
    End If
    m_rs.Close
    
    
    'Cargo el contenido del template en una variable.
    m_contenido = leerArchivo(m_archivo)
    If m_contenido = "" Then
        generarTemplate = ""
        Exit Function
    End If
    
    'FGZ - 08/08/2012 ------------------------------
    AnexoImgHeader = ""
    If Col1 <> 0 Then
        Call CargarArrTagsImg(ArrTags, m_contenido, Col1)
    End If
    'FGZ - 08/08/2012 ------------------------------
    
    
    'LED 18/12/2014 - Busco si el tipo de notificacion es el 18
    StrSql = " SELECT noti_ale.notinro, tnotinro FROM noti_ale " & _
             " INNER JOIN notificacion ON notificacion.notinro = noti_ale.notinro " & _
             " WHERE aleNro = " & Origen
    OpenRecordset StrSql, rs_tiponoti
    If Not rs_tiponoti.EOF Then
        Select Case CLng(rs_tiponoti!TnotiNro)
            Case 18
                m_contenido = reemplazarTagsNotificacion18(m_contenido, ArrTags, rs_tiponoti!notinro)
            Case Else
                m_contenido = reemplazarTags(m_contenido, ArrTags)
        End Select
    Else
        m_contenido = reemplazarTags(m_contenido, ArrTags)
    End If
    'reemplazo los tags que vinieron por ArrTags
    
    
    
    'Realizo al busqueda inversa, busco los tags en el template y despues busco el valor en la bd.
    m_contenido = buscarReemplazarRex(m_contenido)
    
    generarTemplate = GenerarArchivoTemplate(dirSistema, codproc, m_contenido, fileName)
    
    
    Set rs_tiponoti = Nothing
    Set m_rs = Nothing

End Function

Public Sub CargarArrTagsImg(ArrTags, ByVal m_contenido As String, ByVal Ternro As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que busca la imagen del tipo y tercero pasado por parametros y agrega los tags al array.
'               Como efecto colateral llena una lista de adjuntos para el header
' Autor      : FGZ
' Fecha      : 08/08/2012
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Imagen As String
Dim ImagenFull As String
Dim I As Long
Dim m_dir As String
Dim Ok As Boolean

    Call BuscarDirectorioImg(m_dir, Ok)
    
    If Ok Then
        '

        'Imagenes reemplazar nnI00x por imagen de tipo x asociada al tercero pasado como parametro
        'Ademas debe agregar a la lista del header
        For I = 1 To MaxTipoImg
            If InStr(1, m_contenido, "{ññI" & Format(I, "000") & "}") > 0 Then
                'Imagen = BuscarImg(I, Ternro)
                Call BuscarImg(I, Ternro, Imagen, ImagenFull)
                If Not EsNulo(Imagen) Then
                    ArrTags = ArrTags & "{ññI" & Format(I, "000") & "}" & septptepar & Imagen & septptegrp
                    If EsNulo(AnexoImgHeader) Then
                        'AnexoImgHeader = ImagenFull & ";"
                        AnexoImgHeader = m_dir & Imagen & ";"
                    Else
                        'AnexoImgHeader = AnexoImgHeader & ImagenFull & ";"
                        AnexoImgHeader = AnexoImgHeader & m_dir & Imagen & ";"
                    End If
                End If
            End If
        Next I
        'FGZ - 31/07/2012 -------------------------------
    End If
End Sub

Function generarMailHeader(tipoOrigen, Origen, funcion) As String
Dim Ok As Boolean


    'Agrega la ruta a todos los archivos dentro de la carpeta del template
    If (tipoOrigen = "" Or Origen = "" Or funcion = "") Then
        generarMailHeader = ""
        Exit Function
    End If

    Dim m_sql As String
    Dim m_rs As New ADODB.Recordset

    Dim m_archivo As String
    Dim m_dirSistema As String
    Dim m_dirTemplate As String
    Dim m_fileTemplate As String
    Dim m_contenido As String
    
    
    BuscarDirectorioT m_dirSistema, Ok
    
    If Ok = False Then
        generarMailHeader = ""
        Exit Function
    End If

    
    'Busco la carpeta y el archivo del template
    m_sql = " SELECT tpltefile, tpltefolder "
    m_sql = m_sql & " FROM tpte_files "
    m_sql = m_sql & " WHERE ttorigennro = " & tipoOrigen
    m_sql = m_sql & " AND tplteorigen = " & Origen
    m_sql = m_sql & " AND tfunnro = " & funcion
    OpenRecordset m_sql, m_rs
    If Not m_rs.EOF Then
        m_dirTemplate = m_rs("tpltefolder")
        m_fileTemplate = m_rs("tpltefile")
        generarMailHeader = listaArchivos(m_dirSistema & "\" & m_dirTemplate, m_fileTemplate) & AnexoImgHeader
    Else
        generarMailHeader = "" & AnexoImgHeader
    End If
    
    m_rs.Close
    
    Set m_rs = Nothing
    
End Function

Function leerArchivo(archivo) As String
    'Lee el archivo y carga el contenido en una variable (lo devuelve).
    Dim m_fs
    Dim m_TextStream
    Dim m_FileContents As String
    'Creo los objs
    Set m_fs = CreateObject("Scripting.FileSystemObject")
    If m_fs.FileExists(archivo) = True Then
        Set m_TextStream = m_fs.OpenTextFile(archivo, 1) '1=ForReading - 2=ForWriting - 3=ForAppending
        m_FileContents = m_TextStream.ReadAll
        m_TextStream.Close
        Set m_TextStream = Nothing
    Else
        'error
        m_FileContents = ""
    End If
    Set m_fs = Nothing
    leerArchivo = m_FileContents
End Function

Function listaArchivos(ruta, archivo) As String
    Dim m_objFSO
    Dim m_objItem
    Set m_objFSO = CreateObject("Scripting.FileSystemObject")
    Dim m_objFolder
    Set m_objFolder = m_objFSO.GetFolder(ruta)
    listaArchivos = ""
    For Each m_objItem In m_objFolder.Files
        'FGZ - 03/06/2014 ----------------------
        'If m_objItem.Name <> archivo Then
        If m_objItem.Name <> archivo And m_objItem.Name <> "Thumbs.db" Then
            listaArchivos = listaArchivos & ruta & "\" & m_objItem.Name & ";"
        End If
    Next
    Set m_objItem = Nothing
    Set m_objFolder = Nothing
    Set m_objFSO = Nothing
    'listaArchivos = listaArchivos
End Function

Function reemplazarTags(Texto As String, ByVal ArrTags As String) As String
    Dim m_parTags
    Dim m_etiquetaValor
    Dim m_i As Integer
    
    m_parTags = Split(ArrTags, septptegrp) '@@
    For m_i = 0 To UBound(m_parTags)
        m_etiquetaValor = Split(m_parTags(m_i), septptepar)   '::
        If UBound(m_etiquetaValor) > 0 Then
            Texto = Replace(Texto, m_etiquetaValor(0), m_etiquetaValor(1))
        End If
    Next
    reemplazarTags = Texto
End Function

Function reemplazarTagsNotificacion18(ByVal Texto As String, ByVal ArrTags As String, ByVal notinro As Integer) As String
    Dim m_parTags
    Dim m_etiquetaValor
    Dim m_i As Integer
    Dim corte As Integer
    Dim agrupa As Integer
    Dim notidesde As Integer
    Dim notihasta As Integer
    Dim rsAux As New ADODB.Recordset
    Dim seccionRepetitiva As String
    Dim seccionRepetitivaAux As String
    Dim seccionRepetitivaAux2 As String
    Dim repeticiones As String
    Dim reemplazar As Boolean
    Dim prueba As Integer
    
    StrSql = " SELECT colcorte, colagrupa, coldesde, colhasta FROM noti_agrupa WHERE notinro = " & notinro
    OpenRecordset StrSql, rsAux
    If Not rsAux.EOF Then
        corte = rsAux!colcorte
        agrupa = rsAux!colagrupa
        notidesde = rsAux!coldesde
        notihasta = rsAux!colhasta
        prueba = notihasta - agrupa
    End If
    seccionRepetitiva = Mid(Texto, InStr(Texto, "{rep}") + 5, InStr(1, Texto, "{/rep}") - (InStr(Texto, "{rep}") + 5))
    m_parTags = Split(ArrTags, septptegrp) '@@
    repeticiones = 0
    'reemplazar = False
    seccionRepetitivaAux = seccionRepetitiva
    For m_i = 0 To UBound(m_parTags)
        m_etiquetaValor = Split(m_parTags(m_i), septptepar)   '::
        If UBound(m_etiquetaValor) > 0 Then
            
            If Right(Left(m_etiquetaValor(0), 4), 3) = "ññC" Then
                'controlo que sea una columna
                'If CLng(Left(Right(m_etiquetaValor(0), 4), 3)) >= notidesde And CLng(Left(Right(m_etiquetaValor(0), 4), 3)) <= notihasta Then
                'If InStr(seccionRepetitiva, m_etiquetaValor(0)) Then
                    'controlo que sea el primer reemplazo
                
                    If repeticiones <= prueba Then
                        Texto = Replace(Texto, m_etiquetaValor(0), m_etiquetaValor(1))
                    Else
                        seccionRepetitivaAux = Replace(seccionRepetitivaAux, Left(m_etiquetaValor(0), 4) & Right("000" & (((repeticiones) Mod (prueba + 1))), 3) & "}", m_etiquetaValor(1))
                        'reemplazar = True
                        If ((repeticiones + 1) Mod (notihasta - 1) = 0) And (repeticiones - 1 <> prueba) Then
                        'If ((repeticiones - 1) Mod prueba = 0) And (repeticiones - 1 <> prueba) Then
                            seccionRepetitivaAux2 = seccionRepetitivaAux2 & seccionRepetitivaAux & vbCrLf
                            seccionRepetitivaAux = seccionRepetitiva
                        End If
                    End If
                    '& vbCrLf

                    repeticiones = repeticiones + 1

                'Else
                    'Texto = Replace(Texto, m_etiquetaValor(0), m_etiquetaValor(1))
                'End If
            Else
                Texto = Replace(Texto, m_etiquetaValor(0), m_etiquetaValor(1))
            End If
        End If
    Next
    Texto = Replace(Texto, "{rep}", "")
    Texto = Replace(Texto, "{/rep}", "")
    Texto = Replace(Texto, "{rep001}", seccionRepetitivaAux2)
    reemplazarTagsNotificacion18 = Texto
    Set rsAux = Nothing
End Function

Function buscarReemplazarRex(Texto As String) As String
    'Dim StringToSearch
    Dim RexObj
    Dim rexmatch
    Dim rexmatched
    Dim m_valor As String
    
    Set RexObj = New RegExp
    RexObj.Pattern = m_tagrex
    RexObj.IgnoreCase = True
    RexObj.Global = True
    
    Set rexmatch = RexObj.Execute(Texto)
    If rexmatch.Count > 0 Then
        For Each rexmatched In rexmatch
            'rexmatched.Value
            'rexmatched.FirstIndex
            m_valor = buscarEtiquetaBD(rexmatched.Value)
            Texto = RexObj.Replace(Texto, m_valor)
        Next
    End If
    Set RexObj = Nothing
    buscarReemplazarRex = Texto
End Function

Function buscarEtiquetaBD(Etiqueta As String) As String
    'Ver calculo y param
    Dim m_sql As String
    Dim m_rs As New ADODB.Recordset
    Dim m_rs2 As New ADODB.Recordset
    
    m_sql = " SELECT tagcalculo, tagvalor, tagparam "
    m_sql = m_sql & " FROM tpte_tag "
    m_sql = m_sql & " WHERE tagdesabr = '" & Etiqueta & "'"
    OpenRecordset m_sql, m_rs
    If Not m_rs.EOF Then
        Select Case CLng(m_rs("tagcalculo"))
            Case 0:
                buscarEtiquetaBD = m_rs("tagvalor")
            Case 1:
                ' seamo fori, se resuelve local al asp
            Case 2:
                m_sql = m_rs("tagvalor")
                OpenRecordset m_sql, m_rs2
                If Not m_rs.EOF Then
                    buscarEtiquetaBD = m_rs2(0)
                Else
                    buscarEtiquetaBD = ""
                End If
            Case 3:
                ' a definir segun etiquetas.
        End Select
    Else
        buscarEtiquetaBD = Etiqueta
    End If
    m_rs.Close
    Set m_rs = Nothing
End Function

Function GenerarArchivoTemplate(ByVal ruta As String, ByVal codproc As Long, contenido As String, fileName) As String
    Dim m_fs
    Dim m_arch
    Dim m_archpath As String
    Dim m_nombrearch As String
        
    Set m_fs = CreateObject("Scripting.FileSystemObject")

    m_nombrearch = fileName & "_msg_body.html"
    'm_nombrearch = "msg_body_" & codproc & ".html" '"_" & genGuid &
    'm_archpath = ruta       '& "\attach"    'Server.MapPath(ruta) '"/rhprox2/in-out/attach/"
    'm_archpath = replace(m_archpath,"\","\\")
    'm_archpath = m_archpath & "\" & m_nombrearch
    m_archpath = m_nombrearch
    'response.write(ruta & "<br>")
    'response.write(m_archpath)
    'response.end()
    
    Set m_arch = m_fs.CreateTextFile(m_archpath, True)
    m_arch.Write (contenido)
    m_arch.Close
    
    Set m_arch = Nothing
    Set m_fs = Nothing
    
    GenerarArchivoTemplate = m_archpath

End Function



' __________________________________________________________________________________________
' Descripcion: Busca directorio donde se encuentran los templates
' Autor      : Leticia A.
' Fecha      : 10/06/2011
' Ult. Mod   :
' __________________________________________________________________________________________
Public Sub BuscarDirectorioT(ByRef m_dirSistema As String, ByRef Ok As Boolean)
Dim rs As New ADODB.Recordset

On Error GoTo E_BuscarDir

Ok = True


    'Busco el dir de los templates --> No se usa mas
    'm_sql = " SELECT sisdirvirtemplates from sistema "
    'OpenRecordset m_sql, m_rs
    'If Not m_rs.EOF Then
    '    m_dirSistema = m_rs("sisdirvirtemplates")
    'Else
    '    generarTemplate = ""
    '    Exit Function
    'End If
    'm_rs.Close
    
    
    StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro=1 "
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        m_dirSistema = Trim(rs!sis_direntradas)
        m_dirSistema = m_dirSistema & "\Templates"
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el registro - Dir. Entradas - de la tabla sistema nro 1"
        Ok = False
        Exit Sub
    End If
    rs.Close

   
Exit Sub

E_BuscarDir:
    HuboError = True
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: BuscarDirectorioT"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    Ok = False
    
End Sub

Public Sub BuscarDirectorioImg(ByRef m_dirSistema As String, ByRef Ok As Boolean)
' __________________________________________________________________________________________
' Descripcion: Busca directorio donde se encuentran las imagenes
' Autor      : FGZ
' Fecha      : 08/08/2012
' Ult. Mod   :
' __________________________________________________________________________________________
Dim rs As New ADODB.Recordset
On Error GoTo E_BuscarDir

    Ok = True

    StrSql = "SELECT sis_dirimag FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        If Not EsNulo(Trim(rs!sis_dirimag)) Then
            m_dirSistema = Trim(rs!sis_dirimag)
            If Right(m_dirSistema, 1) <> "\" Then
                m_dirSistema = m_dirSistema & "\"
            End If
        Else
            Flog.writeline Espacios(Tabulador * 1) & "El Directorio de Imagenes (sis_dirimag) - de la tabla sistema no es valido"
            Ok = False
            Exit Sub
        End If
        
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el registro - Dir. Imagen - de la tabla sistema nro 1"
        Ok = False
        Exit Sub
    End If
    rs.Close
  
Exit Sub

E_BuscarDir:
    HuboError = True
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: BuscarDirectorioT"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    Ok = False
End Sub

