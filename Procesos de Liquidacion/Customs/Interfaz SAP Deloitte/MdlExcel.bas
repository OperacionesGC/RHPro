Attribute VB_Name = "MdlExcel"
Option Explicit

Public ExcelSheet As New Excel.Application

'--------------------------------------------------------------
'--------------------------------------------------------------
' Para utilizarlo, agregar la siguiente referencia:
'   Project-> References-> Select MS Excel 10.0 Object Library
'--------------------------------------------------------------
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
' Fecha      : 27/03/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim i As Integer
Dim NombreHoja As String
Dim SeUsa As Boolean

On Error GoTo ME_Local

    Set ExcelSheet = New Excel.Application
    ExcelSheet.SheetsInNewWorkbook = 2
    ExcelSheet.Workbooks.Add
    ExcelSheet.Caption = nombre
    
    
    NombreHoja = "Formato1" 'ABM
    ExcelSheet.Sheets(1).Name = NombreHoja
    Call CargarColumnas(1, NombreHoja)
    
    NombreHoja = "Formato2" 'Infotipos
    ExcelSheet.Sheets(2).Name = NombreHoja
    Call CargarColumnas(2, NombreHoja)
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

On Error GoTo ME_Local
    Select Case Hoja
    Case 1: 'ABM
        Call Columnas_ABM(Hoja)
    Case 2: 'Infotipos
        Call Columnas_Infotipo(Hoja)
    End Select

ME_Local:
    Flog.writeline Err.Description
End Sub


Public Sub Columnas_Infotipo(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 27/03/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

    i = 1
    J = 1

    NombreColumna = "Numero de Person."
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Infotipo"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Descripcion"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Importe"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna

    'Doy formato a la hoja
    Maximo = J
    For J = 1 To Maximo
        ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
        Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
        Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
    Next J
End Sub

Public Sub Columnas_ABM(ByVal Hoja As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 27/03/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim J As Integer
Dim i As Integer
Dim Maximo As Integer
Dim NombreColumna As String

    i = 1
    J = 1

    NombreColumna = "Numero de Person."
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Tipo de ID"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Descripcion"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Numero de Identidad"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Apellido"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "2do Apellido"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Nombre"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Cuil"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Genero"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Fecha de Nac."
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Estado Civil"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Nacionalidad"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Division de Personal"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Descripcion"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Grupo de Personal"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Area de Personal"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Area Funcional"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Tipo de Medida"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Motivo Medida"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Status del empleado"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Codigo de Compañia"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Descripcion"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Subdivision de Personal"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Descripcion"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Centro de Costo"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Descripcion"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Area Nomina"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Relacion Laboral"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Descripcion"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Posicion"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Descripcion"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Funcion"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Descripcion"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Unidad Organizativa"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Descripcion"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Tipo de Contrato"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Descripcion"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Vencimiento"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Periodo de Prueba #"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Periodo de Prueba (UN)"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Regla plan horario de trabajo"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Descripcion"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Porcentaje de Horario de trabajo"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Clase de Datos Bancarios"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Descripcion"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Clave de Banco"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Banco"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Cuenta Bancaria"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Clv. Ctrl. Bancos"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Via de Pago"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Sucursal"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Fecha de Medida"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Fecha Inicio"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Fecha Fin"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Obra Social"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Descripcion"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Plan OS"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Capitalizacion"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Reparto"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Actividad SIJP"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Codicion SIJP"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Descripcion"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Tipo de Servicios"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Calle"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Nro"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Piso"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Depto"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "CP"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Localidad"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna
    J = J + 1
    NombreColumna = "Provincia"
    ExcelSheet.Sheets(Hoja).Cells(i, J) = NombreColumna


    'Doy formato a la hoja
    Maximo = J
    For J = 1 To Maximo
        ExcelSheet.Sheets(Hoja).Cells(i, J).Interior.Color = vbCyan
        Sheets(Hoja).Cells(i, J).Interior.Pattern = xlSolid
        Sheets(Hoja).Cells(i, J).BorderAround 1, xlThin, xlColorIndexAutomatic
    Next J
End Sub


Public Sub Insertar_Valor_Excel(ByVal Hoja As Integer, ByVal Fila As Long, ByVal Columna As Long, ByVal valor As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga los nombres de las columnas de cada hoja.
' Autor      : FGZ
' Fecha      : 06/07/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

    On Error GoTo ME_Local
    
    Columna = Columna + 1
    valor = Replace(valor, "+", "")

    ExcelSheet.Sheets(Hoja).Cells(Fila, Columna).NumberFormat = "@"
    ExcelSheet.Sheets(Hoja).Cells(Fila, Columna) = valor
Exit Sub
ME_Local:
    Flog.writeline Err.Description
End Sub
