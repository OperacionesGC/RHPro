Attribute VB_Name = "MdlExcel"
Option Explicit

Public ExcelSheet As New Excel.Application

'--------------------------------------------------------------
' Para utilizarlo, agregar la siguiente referencia:
'   Project-> References-> Select MS Excel 10.0 Object Library
'--------------------------------------------------------------

Public Sub CerrarArchivoExcel(ByVal nombre As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Cierra un libro de Excel.
' Autor      : FGZ
' Fecha      : 07/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
On Error GoTo ME_Local

    ExcelSheet.ActiveWorkbook.SaveAs FileName:=nombre, FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False

    ExcelSheet.Application.Quit
    ExcelSheet.Workbooks.Close
    Set ExcelSheet = Nothing
    
Exit Sub
ME_Local:
    Flog.writeline Err.Description
End Sub

Public Function CrearArchivoExcel(ByVal nombre As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Crea un libro de Excel.
' Autor      : FGZ
' Fecha      : 05/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim i As Integer
Dim NombreHoja As String
Dim SeUsa As Boolean

On Error GoTo ME_Local

    Set ExcelSheet = New Excel.Application
    ExcelSheet.SheetsInNewWorkbook = 22
    ExcelSheet.Workbooks.Add
    ExcelSheet.Caption = nombre

    For i = 1 To 22
        Select Case i
        Case 1:
            NombreHoja = "IT0000"   '- ACTIONS
        Case 2:
            NombreHoja = "IT0001"   '- ORGANIZATIONAL ASSIGNMENT
        Case 3:
            NombreHoja = "IT0002"   '- PERSONAL DATA
        Case 4:
            NombreHoja = "IT0006"   '- ADDRESSES
        Case 5:
            NombreHoja = "IT0008"   '– BASIC PAY
        Case 6:
            NombreHoja = "IT0009"   '– BANK DETAILS
        Case 7:
            NombreHoja = "IT0014"   '– DEVENGOS Y DEDUCCIONES PERIODICAS
        Case 8:
            NombreHoja = "IT0015"   '– DEVENGOS COMPLEMENTARIOS
        Case 9:
            NombreHoja = "IT0021"   '- FAMILY/RELEATED PERSON
        Case 10:
            NombreHoja = "IT0041"   '– DATE SPECIFICATIONS
        Case 11:
            NombreHoja = "IT0057"   '– MEMBERSHIP FEES
        Case 12:
            NombreHoja = "IT0185"   '- PERSONAL IDs
        Case 13:
            NombreHoja = "IT0389"   '- IMPUESTO A LAS GANANCIAS (ARGENTINA)
        Case 14:
            NombreHoja = "IT0390"   '– IMPUESTO A LAS GANANCIAS - DEDUCCIONES (ARGENTINA)
        Case 15:
            NombreHoja = "IT0391"   '– IMPUESTO A LAS GANANCIAS - OTRO EMPLEADOR
        Case 16:
            NombreHoja = "IT0392"   '– SEGURIDAD SOCIAL - ARGENTINA
        Case 17:
            NombreHoja = "IT0393"   '– DATOS DE FAMILIA AYUDA ESCOLAR (ARGENTINA)
        Case 18:
            NombreHoja = "IT0394"   '– DATOS FAMILIA: INFORMACION ADICIONAL -ARGENTINA
        Case 19:
            NombreHoja = "IT2001"   '– AUSENTISMOS
        Case 20:
            NombreHoja = "IT2010"   '– COMPROBANTE DE REMUNERACION
        Case 21:
            NombreHoja = "IT9004"   '– COMPROBANTE DE REMUNERACION
        Case 22:
            NombreHoja = "IT9302"   '– Prestamos
        End Select
        ExcelSheet.Sheets(i).Name = NombreHoja
        Call CargarColumnas(i, NombreHoja)
    Next i
    
Exit Function

ME_Local:
    Flog.writeline Err.Description
End Function


Public Sub CargarColumnas(ByVal Hoja As Integer, ByVal Infotipo As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 05/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

        Select Case Infotipo
        Case "IT0000":  '- ACTIONS
            Call Columnas_Infotipo0000(Hoja)
        Case "IT0001":  '– ORGANIZATIONAL ASSIGNMENT
            Call Columnas_Infotipo0001(Hoja)
        Case "IT0002":  '- PERSONAL DATA
            Call Columnas_Infotipo0002(Hoja)
        Case "IT0006":  '– ADDRESSES
            Call Columnas_Infotipo0006(Hoja)
        Case "IT0008":  '– BASIC PAY
            Call Columnas_Infotipo0008(Hoja)
        Case "IT0009":  '– BANK DETAILS
            Call Columnas_Infotipo0009(Hoja)
        Case "IT0014":  '– DEVENGOS Y DEDUCCIONES PERIODICAS
            Call Columnas_Infotipo0014(Hoja)
        Case "IT0015":  '– DEVENGOS COMPLEMENTARIOS
            Call Columnas_Infotipo0015(Hoja)
        Case "IT0021":  '- FAMILY/RELEATED PERSON
            Call Columnas_Infotipo0021(Hoja)
'        Case "IT0027":  '- Distribucion de Costos
'            Call Columnas_Infotipo0027(Hoja)
        Case "IT0041":  '– DATE SPECIFICATIONS
            Call Columnas_Infotipo0041(Hoja)
        Case "IT0057":  '– MEMBERSHIP FEES
            Call Columnas_Infotipo0057(Hoja)
        Case "IT0185":  '- PERSONAL IDs
            Call Columnas_Infotipo0185(Hoja)
        Case "IT0389":  '- IMPUESTO A LAS GANANCIAS (ARGENTINA)
            Call Columnas_Infotipo0389(Hoja)
        Case "IT0390":  '– IMPUESTO A LAS GANANCIAS - DEDUCCIONES (ARGENTINA)
            Call Columnas_Infotipo0390(Hoja)
        Case "IT0391":  '– IMPUESTO A LAS GANANCIAS - OTRO EMPLEADOR
            Call Columnas_Infotipo0391(Hoja)
        Case "IT0392":  '– SEGURIDAD SOCIAL - ARGENTINA
            Call Columnas_Infotipo0392(Hoja)
        Case "IT0393":  '– DATOS DE FAMILIA AYUDA ESCOLAR (ARGENTINA)
            Call Columnas_Infotipo0393(Hoja)
        Case "IT0394":  '– DATOS FAMILIA: INFORMACION ADICIONAL -ARGENTINA
            Call Columnas_Infotipo0394(Hoja)
        Case "IT2001":  '– AUSENTISMOS
            Call Columnas_Infotipo2001(Hoja)
        Case "IT2010":  '– COMPROBANTE DE REMUNERACION
            Call Columnas_Infotipo2010(Hoja)
        Case "IT9004":  '–
            Call Columnas_Infotipo9004(Hoja)
        Case "IT9302":  '– Prestamos
            Call Columnas_Infotipo9302(Hoja)
        End Select
End Sub

Public Sub Columnas_Infotipo0000(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 05/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

            i = 1
            J = 1
            NombreColumna = "PERNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "INFTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "SUBTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BEGDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ENDDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PREAS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "MASSN"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "MASSG"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "STAT1"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "STAT2"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "STAT3"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            
            'Doy formato a la hoja
            Maximo = J
            For J = 1 To Maximo
                ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
                Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
                Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
            Next J

End Sub


Public Sub Columnas_Infotipo0001(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 05/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

    i = 1
    J = 1
    NombreColumna = "PERNR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "INFTY"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "SUBTY"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BEGDA"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ENDDA"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BUKRS"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "WERKS"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PERSG"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PERSK"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VDSK1"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GSBER"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BTRTL"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "JUPER"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ABKRS"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANSVH"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "KOSTL"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ORGEH"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PLANS"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "STELL"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "MSTBR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "SACHA"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "SACHP"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "SACHZ"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "SNAME"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ENAME"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OTYPE"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "SBMOD"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "KOKRS"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FISTL"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GEBER"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            
    'Doy formato a la hoja
    Maximo = J
    For J = 1 To Maximo
        ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
        Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
        Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
    Next J
End Sub

Public Sub Columnas_Infotipo0002(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 05/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

    i = 1
    J = 1

    NombreColumna = "PERNR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "INFTY"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "SUBTY"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BEGDA"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ENDDA"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "INITS"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NACHN"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NAME2"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NACH2"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VORNA"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CNAME"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TITEL"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TITL2"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NAMZU"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VORSW"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VORS2"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "RUFNM"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "MIDNM"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "KNZNM"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANRED"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GESCH"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GBDAT"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GBLND"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GBDEP"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GBORT"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NATIO"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NATI2"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NATI3"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "SPRSL"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "KONFE"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FAMST"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FAMDT"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANZKD"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NACON"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PERMO"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PERID"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GBPAS"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FNAMK"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LNAMK"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FNAMR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LNAMR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NABIK"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NABIR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NICKK"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NICKR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GBJHR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GBMON"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GBTAG"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NCHMC"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VNAMC"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NAMZ2"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            
    'Doy formato a la hoja
    Maximo = J
    For J = 1 To Maximo
        ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
        Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
        Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
    Next J
End Sub


Public Sub Columnas_Infotipo0006(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 05/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

    i = 1
    J = 1

    NombreColumna = "PERNR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "INFTY"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "SUBTY"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BEGDA"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ENDDA"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANSSA"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NAME2"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "STRAS"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ORT01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ORT02"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PSTLZ"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LAND1"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TELNR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ENTKM"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "WKWNG"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BUSRT"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LOCAT"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ADR03"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ADR04"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "STATE"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "HSNMR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "POSTA"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BLDNG"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FLOOR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "STRDS"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ENTK2"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "COM01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NUM01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "COM02"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NUM02"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "COM03"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NUM03"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "COM04"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NUM04"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "COM05"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NUM05"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "COM06"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NUM06"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "INDRL"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "COUNC"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "RCTVC"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OR2KK"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CONKK"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OR1KK"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "RAILW"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            
    'Doy formato a la hoja
    Maximo = J
    For J = 1 To Maximo
        ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
        Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
        Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
        
    Next J
End Sub


Public Sub Columnas_Infotipo0008(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 05/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

    i = 1
    J = 1

    NombreColumna = "PERNR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "INFTY"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "SUBTY"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BEGDA"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ENDDA"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PREAS"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TRFAR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TRFGB"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TRFGR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TRFST"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "STVOR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ORZST"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PARTN"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "WAERS"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VGLTA"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VGLGB"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VGLGR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VGLST"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VGLSV"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BSGRD"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "DIVGV"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANSAL"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FALGK"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FALGR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LGA01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BET01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANZ01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "EIN01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OPK01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LGA02"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BET02"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANZ02"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "EIN02"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OPK02"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LGA03"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BET03"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANZ03"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "EIN03"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OPK03"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LGA04"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BET04"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANZ04"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "EIN04"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OPK04"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LGA05"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BET05"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANZ05"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "EIN05"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OPK05"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LGA06"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BET06"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANZ06"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "EIN06"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OPK06"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LGA07"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BET07"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANZ07"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "EIN07"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OPK07"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LGA08"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BET08"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANZ08"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "EIN08"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OPK08"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LGA09"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BET09"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANZ09"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "EIN09"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OPK09"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LGA10"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BET10"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANZ10"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "EIN10"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OPK10"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LGA11"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BET11"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANZ11"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "EIN11"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OPK11"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LGA12"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BET12"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANZ12"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "EIN12"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OPK12"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LGA13"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BET13"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANZ13"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "EIN13"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OPK13"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LGA14"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BET14"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANZ14"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "EIN14"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OPK14"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LGA15"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BET15"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANZ15"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "EIN15"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OPK15"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LGA16"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BET16"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANZ16"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "EIN16"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OPK16"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LGA17"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BET17"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANZ17"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "EIN17"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OPK17"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LGA18"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BET18"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANZ18"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "EIN18"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OPK18"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LGA19"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BET19"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANZ19"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "EIN19"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OPK19"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LGA20"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BET20"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANZ20"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "EIN20"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "OPK20"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "IND01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "IND02"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "IND03"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "IND04"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "IND05"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "IND06"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "IND07"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "IND08"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "IND09"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "IND10"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "IND11"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "IND12"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "IND13"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "IND14"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "IND15"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "IND16"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "IND17"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "IND18"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "IND19"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "IND20"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ANCUR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CPIND"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna

    'Doy formato a la hoja
    Maximo = J
    For J = 1 To Maximo
        ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
        Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
        Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
    Next J
End Sub

Public Sub Columnas_Infotipo0009(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 05/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

            
            i = 1
            J = 1

            NombreColumna = "PERNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "INFTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "SUBTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BEGDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ENDDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "OPKEN"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BETRG"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "WAERS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ANZHL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ZEINH"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BNKSA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ZLSCH"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "EMFTX"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BKPLZ"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BKORT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BANKS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BANKL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BANKN"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BANKP"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BKONT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "SWIFT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DTAWS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DTAMS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "STCD1"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "STCD2"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PSKTO"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ESRNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ESRRE"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ESRPZ"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "EMFSL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ZWECK"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BTTYP"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PAYTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PAYID"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "OCRSN"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BONDT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BKREF"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "STRAS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "State"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna

            'Doy formato a la hoja
            Maximo = J
            For J = 1 To Maximo
                ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
                Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
                Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
            Next J
End Sub

Public Sub Columnas_Infotipo0014(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 05/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

            
            i = 1
            J = 1

            NombreColumna = "PERNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "INFTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "SUBTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BEGDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ENDDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGART"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "OPKEN"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BETRG"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "WAERS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ANZHL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ZEINH"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "INDBW"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ZDATE"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ZFPER"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ZANZL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ZEINZ"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ZUORD"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "UWDAT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "MODEL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PREAS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna

            'Doy formato a la hoja
            Maximo = J
            For J = 1 To Maximo
                ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
                Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
                Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
            Next J
End Sub


Public Sub Columnas_Infotipo0015(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 05/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

            i = 1
            J = 1

            NombreColumna = "PERNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "INFTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "SUBTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BEGDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ENDDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGART"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "OPKEN"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BETRG"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "WAERS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ANZHL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ZEINH"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "INDBW"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ZUORD"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ESTDT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PABRJ"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PABRP"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "UWDAT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ITFTT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PREAS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna

            'Doy formato a la hoja
            Maximo = J
            For J = 1 To Maximo
                ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
                Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
                Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
            Next J
End Sub


Public Sub Columnas_Infotipo0021(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 05/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

            
            i = 1
            J = 1

            NombreColumna = "PERNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "INFTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "SUBTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BEGDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ENDDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "OBJPS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FAMSA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FGBDT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FGBLD"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FANAT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FASEX"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FAVOR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FANAM"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FGBOT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FGDEP"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ERBNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FGBNA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FNAC2"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FCNAM"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FKNZN"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FINIT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FVRSW"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FVRS2"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FNMZU"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "AHVNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "KDSVH"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "KDBSL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "KDUTB"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "KDGBR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "KDART"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "KDZUG"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "KDZUL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "KDVBE"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ERMNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "AUSVL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "AUSVG"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FASDT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FASAR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FASIN"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "EGAGA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FANA2"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FANA3"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BETRG"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "TITEL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "EMRGN"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna

            'Doy formato a la hoja
            Maximo = J
            For J = 1 To Maximo
                ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
                Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
                Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
            Next J

End Sub

'Public Sub Columnas_Infotipo0027(ByVal Hoja As Integer)
'' ---------------------------------------------------------------------------------------------
'' Descripcion: Carga los nombres de las columnas de cada hoja.
'' Autor      : FGZ
'' Fecha      : 05/07/2005
'' Ultima Mod.:
'' Descripcion:
'' ---------------------------------------------------------------------------------------------
'Dim J As Integer
'Dim I As Integer
'Dim Maximo As Integer
'Dim NombreColumna As String
'
'
'            I = 1
'            J = 1
'
'            NombreColumna = "PERNR"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "INFTY"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "SUBTY"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "BEGDA"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "ENDDA"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KSTAR"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU01"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB01"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST01"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR01"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU02"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB02"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST02"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR02"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU03"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB03"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST03"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR03"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU04"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB04"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST04"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR04"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU05"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB05"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST05"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR05"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU06"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB06"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST06"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR06"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU07"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB07"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST07"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR07"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU08"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB08"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST08"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR08"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU09"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB09"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST09"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR09"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU10"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB10"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST10"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR10"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU11"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB11"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST11"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR11"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU12"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB12"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST12"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR12"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU13"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB13"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST13"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR13"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU14"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB14"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST14"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR14"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU15"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB15"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST15"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR15"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU16"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB16"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST16"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR16"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU17"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB17"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST17"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR17"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU18"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB18"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST18"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR18"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU19"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB19"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST19"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR19"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU20"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB20"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST20"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR20"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU21"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB21"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST21"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR21"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU22"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB22"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST22"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR22"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU23"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB23"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST23"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR23"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU24"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB24"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST24"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR24"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KBU25"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KGB25"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KST25"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "KPR25"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "FCT01"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "FCD01"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "FCT02"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "FCD02"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "FCT03"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "FCD03"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "FCT04"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "FCD04"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "FCT05"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "FCD05"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "FCT06"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "FCD06"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "FCT07"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "FCD07"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "FCT08"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "FCD08"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF01"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF02"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF03"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF04"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF05"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF06"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF07"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF08"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF09"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF10"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF11"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF12"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF13"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF14"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF15"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF16"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF17"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF18"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF19"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF20"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF21"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF22"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF23"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF24"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "AUF25"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP01"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP02"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP03"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP04"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP05"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP06"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP07"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP08"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP09"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP10"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP11"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP12"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP13"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP14"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP15"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP16"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP17"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP18"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP19"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP20"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP21"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP22"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP23"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP24"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'            J = J + 1
'            NombreColumna = "PSP25"
'            ExcelSheet.Sheets(Hoja).Cells(I, J) = NombreColumna
'
'            'Doy formato a la hoja
'            Maximo = J
'            For J = 1 To Maximo
'                ExcelSheet.Sheets(Hoja).Cells(I, J).Interior.Color = vbCyan
'                Sheets(Hoja).Cells(I, J).Interior.Pattern = xlSolid
'                Sheets(Hoja).Cells(I, J).BorderAround 1, xlThin, xlColorIndexAutomatic
'            Next J
'
'End Sub

Public Sub Columnas_Infotipo0041(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 05/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

            
            i = 1
            J = 1

            NombreColumna = "PERNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "INFTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "SUBTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BEGDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ENDDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAR01"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAT01"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAR02"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAT02"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAR03"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAT03"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAR04"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAT04"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAR05"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAT05"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAR06"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAT06"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAR07"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAT07"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAR08"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAT08"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAR09"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAT09"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAR10"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAT10"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAR11"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAT11"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAR12"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DAT12"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna

            'Doy formato a la hoja
            Maximo = J
            For J = 1 To Maximo
                ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
                Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
                Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
            Next J
End Sub


Public Sub Columnas_Infotipo0057(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 05/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

            i = 1
            J = 1

            NombreColumna = "PERNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "INFTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "SUBTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BEGDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ENDDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "EMFSL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "MTGLN"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "GRPRG"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BETRG"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "WAERS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGART"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ANZHL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ZEINH"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ZFPER"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ZDATE"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ZANZL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ZEINZ"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PRITY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "UFUNC"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "UNLOC"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "USTAT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ESRNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ESRRE"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ESRPZ"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ZWECK"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "OPKEN"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "INDBW"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ZSCHL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "UWDAT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "MODEL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "MGART"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna

            'Doy formato a la hoja
            Maximo = J
            For J = 1 To Maximo
                ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
                Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
                Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
            Next J
End Sub

Public Sub Columnas_Infotipo0185(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 05/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

            
            i = 1
            J = 1

            NombreColumna = "PERNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "INFTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "SUBTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BEGDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ENDDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ICTYP"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ICNUM"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ICOLD"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "AUTH1"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DOCN1"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FPDAT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "EXPID"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ISSPL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ISCOT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IDCOT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "OVCHK"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ASTAT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "AKIND"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "REJEC"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "USEFR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "USETO"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DATEN"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DATEU"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "TIMES"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna

            'Doy formato a la hoja
            Maximo = J
            For J = 1 To Maximo
                ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
                Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
                Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
            Next J
End Sub


Public Sub Columnas_Infotipo0389(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 05/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

            i = 1
            J = 1
            
            NombreColumna = "PERNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "INFTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "SUBTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BEGDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ENDDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IMPUE"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PRNHA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "TOPER"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "TCUIT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DGIOF"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PIMNI"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "CERNO"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "TAEGR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna

            'Doy formato a la hoja
            Maximo = J
            For J = 1 To Maximo
                ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
                Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
                Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
            Next J

End Sub

Public Sub Columnas_Infotipo0390(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 05/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

            i = 1
            J = 1

            NombreColumna = "PERNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "INFTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "SUBTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BEGDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ENDDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGART"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IMDED"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "WAERS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PRMES"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PRFEC"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "NUMIN"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "UNITN"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ICTYP"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ICNUM"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "NOMBR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "STRAS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "HSNMR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "Floor"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "POSTA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PSTLZ"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ORT01"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ORT02"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "State"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LAND1"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna

            'Doy formato a la hoja
            Maximo = J
            For J = 1 To Maximo
                ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
                Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
                Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
            Next J

End Sub

Public Sub Columnas_Infotipo0391(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 05/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

            
            i = 1
            J = 1

            NombreColumna = "PERNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "INFTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "SUBTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BEGDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ENDDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ICTYP"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ICNUM"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "NOMBR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "STRAS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "HSNMR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "Floor"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "POSTA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PSTLZ"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ORT01"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ORT02"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "State"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LAND1"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "WAERS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGA01"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IMP01"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGA02"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IMP02"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGA03"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IMP03"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGA04"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IMP04"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGA05"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IMP05"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGA06"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IMP06"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGA07"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IMP07"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGA08"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IMP08"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGA09"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IMP09"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGA10"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IMP10"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGA11"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IMP11"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGA12"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IMP12"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGA13"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IMP13"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGA14"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IMP14"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGA15"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IMP15"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGA16"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IMP16"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGA17"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IMP17"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGA18"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IMP18"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGA19"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IMP19"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGA20"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "IMP20"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna

            'Doy formato a la hoja
            Maximo = J
            For J = 1 To Maximo
                ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
                Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
                Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
            Next J

End Sub

Public Sub Columnas_Infotipo0392(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 05/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

            
            i = 1
            J = 1

            NombreColumna = "PERNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "INFTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "SUBTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BEGDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ENDDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "OBRAS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "OSNOA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "OBRAO"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "SYJUB"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "CAFJP"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "TYACT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PLANS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "CSERV"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "AFJUB"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ASPCE"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "TPUOS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna

            'Doy formato a la hoja
            Maximo = J
            For J = 1 To Maximo
                ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
                Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
                Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
            Next J

End Sub


Public Sub Columnas_Infotipo0393(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 05/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

            
            i = 1
            J = 1

            NombreColumna = "PERNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "INFTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "SUBTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BEGDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ENDDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FAMSA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "CERAP"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "CERAA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "MESLI"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "OBJPS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna

            'Doy formato a la hoja
            Maximo = J
            For J = 1 To Maximo
                ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
                Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
                Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
            Next J

End Sub

Public Sub Columnas_Infotipo0394(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 05/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

            
            i = 1
            J = 1

            NombreColumna = "PERNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "INFTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "SUBTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BEGDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ENDDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ASFAX"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DISCP"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "TRABA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ESTUD"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FEINF"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "NADOC"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FAMST"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ADHOS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "Clase"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "CCUIL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ICTYP"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ICNUM"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "OBJPS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna

            'Doy formato a la hoja
            Maximo = J
            For J = 1 To Maximo
                ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
                Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
                Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
            Next J
End Sub

Public Sub Columnas_Infotipo2001(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 05/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

            
            i = 1
            J = 1

            NombreColumna = "PERNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "INFTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "SUBTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BEGDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ENDDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BEGUZ"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ENDUZ"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "VTKEN"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "AWART"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ABWTG"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "STDAZ"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ABRTG"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ABRST"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ANRTG"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LFZED"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "KRGED"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "KBBEG"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "RMDDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "KENN1"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "KENN2"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "KALTG"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "URMAN"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BEGVA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BWGRL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "AUFKZ"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "TRFGR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "TRFST"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PRAKN"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PRAKZ"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "OTYPE"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PLANS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "MLDDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "MLDUZ"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "RMDUZ"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "VORGS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "UMSKD"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "UMSCH"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "REFNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "UNFAL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "STKRV"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "STUND"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PSARB"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "AINFT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "GENER"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "HRSIF"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ALLDF"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "WAERS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LOGSYS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "AWTYP"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "AWREF"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "AWORG"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DOCSY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "DOCNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PAYTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PAYID"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BONDT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "OCRSN"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "SPPE1"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "SPPE2"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "SPPE3"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "SPPIN"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ZKMKT"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "FAPRS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "TDLANGU"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "TDSUBLA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "TDTYPE"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna

            'Doy formato a la hoja
            Maximo = J
            For J = 1 To Maximo
                ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
                Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
                Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
            Next J

End Sub

Public Sub Columnas_Infotipo2010(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 05/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

            
            i = 1
            J = 1

            NombreColumna = "PERNR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "INFTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "SUBTY"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BEGDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ENDDA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BEGUZ"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ENDUZ"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "VTKEN"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "STDAZ"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LGART"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ANZHL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ZEINH"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BWGRL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "AUFKZ"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "BETRG"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "ENDOF"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "UFLD1"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "UFLD2"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "UFLD3"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "KEYPR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "TRFGR"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "TRFST"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PRAKN"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PRAKZ"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "OTYPE"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "PLANS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "VERSL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "EXBEL"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "WAERS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "LOGSYS"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "AWTYP"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "AWREF"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "AWORG"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "WTART"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "TDLANGU"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "TDSUBLA"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
            J = J + 1
            NombreColumna = "TDTYPE"
            ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna

            'Doy formato a la hoja
            Maximo = J
            For J = 1 To Maximo
                ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
                Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
                Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
            Next J
End Sub

Public Sub Columnas_Infotipo9004(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 12/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

    i = 1
    J = 1

    NombreColumna = "PERNR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "INFTY"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "SUBTY"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BEGDA"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ENDDA"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TYPE"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BEGDA_T"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ENDDA_T"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "APPLICY"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BEGDA_AP"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ENDDA_AP"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Amount"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BEGDA_AM"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ENDDA_AM"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PERCENT_PLAN"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Status"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BEGDA_PLAN"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ENDDA_PLAN"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "MONTH_SAVIN"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PERCENT_SAVIN"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BEGDA_SAVIN"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ENDDA_SAVIN"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BENE01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "DATE01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PERC01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BENE02"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "DATE02"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PERC02"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BENE03"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "DATE03"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PERC03"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BENE04"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "DATE04"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PERC04"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BENE05"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "DATE05"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PERC05"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BENE06"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "DATE06"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PERC06"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BENE07"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "DATE07"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PERC07"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BENE08"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "DATE08"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PERC08"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BENE09"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "DATE09"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PERC09"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BENE10"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "DATE10"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PERC10"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "RECRE"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TICKET"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "RESTAUR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna

    'Doy formato a la hoja
    Maximo = J
    For J = 1 To Maximo
        ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
        Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
        Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
    Next J
End Sub

Public Sub Columnas_Infotipo9302(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 13/07/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

    i = 1
    J = 1

    NombreColumna = "PERNR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "INFTY"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "SUBTY"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BEGDA"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "ENDDA"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NROPR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TIPPR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Monto"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Moneda"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NOMTES"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "MODLD"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NCUOTAS"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CODIN"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PTAJE"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NPOLIZA"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LAND"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BANKL"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BANKN"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VLRSEG"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NPOLIZA01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "LAND01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BANKL01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BANKN01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VLRSEG01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GARANTIA"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "DETGAR01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VLRGAR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VLRBIEN"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GR_PTVL"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GR_VALOR"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GR_PERDC"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    
    J = J + 1
    NombreColumna = "CPT01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PTV01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VAL01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BAE01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FRE01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GRA01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NCT01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TOT01"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CPT02"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PTV02"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VAL02"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BAE02"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FRE02"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GRA02"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NCT02"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TOT02"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CPT03"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PTV03"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VAL03"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BAE03"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FRE03"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GRA03"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NCT03"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TOT03"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CPT04"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PTV04"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VAL04"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BAE04"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FRE04"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GRA04"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NCT04"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TOT04"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CPT05"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PTV05"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VAL05"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BAE05"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FRE05"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GRA05"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NCT05"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TOT05"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CPT06"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PTV06"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VAL06"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BAE06"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FRE06"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GRA06"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NCT06"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TOT06"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CPT07"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PTV07"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VAL07"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BAE07"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FRE07"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GRA07"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NCT07"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TOT07"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CPT08"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PTV08"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VAL08"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BAE08"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FRE08"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GRA08"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NCT08"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TOT08"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CPT09"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PTV09"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VAL09"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BAE09"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FRE09"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GRA09"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NCT09"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TOT09"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CPT010"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PTV010"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VAL010"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BAE010"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FRE010"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GRA010"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NCT010"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TOT010"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CPT011"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PTV011"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VAL011"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BAE011"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FRE011"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GRA011"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NCT011"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TOT011"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CPT012"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PTV012"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VAL012"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BAE012"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FRE012"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GRA012"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NCT012"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TOT012"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CPT013"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PTV013"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VAL013"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BAE013"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FRE013"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GRA013"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NCT013"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TOT013"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CPT014"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PTV014"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VAL014"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BAE014"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FRE014"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GRA014"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NCT014"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TOT014"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CPT015"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PTV015"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VAL015"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BAE015"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FRE015"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GRA015"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NCT015"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TOT015"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CPT016"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PTV016"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VAL016"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BAE016"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FRE016"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GRA016"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NCT016"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TOT016"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CPT017"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PTV017"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VAL017"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BAE017"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FRE017"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GRA017"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NCT017"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TOT017"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CPT018"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PTV018"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VAL018"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BAE018"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FRE018"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GRA018"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NCT018"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TOT018"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CPT019"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PTV019"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VAL019"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BAE019"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FRE019"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GRA019"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NCT019"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TOT019"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CPT020"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "PTV020"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "VAL020"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "BAE020"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "FRE020"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "GRA020"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "NCT020"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "TOT020"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna

    'Doy formato a la hoja
    Maximo = J
    For J = 1 To Maximo
        ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
        Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
        Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
    Next J
End Sub

Public Sub Insertar_Valor_Excel(ByVal Hoja As Integer, ByVal Fila As Long, ByVal Columna As Long, ByVal Valor As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 06/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

    On Error GoTo ME_Local
    
    Valor = Replace(Valor, "+", "")

    ExcelSheet.Sheets(Hoja).Cells(Fila, Columna).NumberFormat = "@"
    ExcelSheet.Sheets(Hoja).Cells(Fila, Columna) = Valor
Exit Sub
ME_Local:
    Flog.writeline Err.Description
End Sub
