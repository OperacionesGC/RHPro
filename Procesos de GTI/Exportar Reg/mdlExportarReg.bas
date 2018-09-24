Attribute VB_Name = "mdlExportarReg"
Option Explicit
Dim fs
Dim freg

Dim FechaDesde As Date
Dim FechaHasta As Date
Dim Separador As String
Sub Main()
Dim Archivo As String
Dim pos As Integer
Dim strcmdLine  As String
 
 
' carga las configuraciones basicas, formato de fecha, string de conexion,
' tipo de BD y ubicacion del archivo de log
Call CargarConfiguracionesBasicas

 strcmdLine = Command()
 
 pos = InStr(1, strcmdLine, ",") - 1
 FechaDesde = CDate(Left(strcmdLine, pos))
 FechaHasta = CDate(Mid(strcmdLine, pos + 2))
    
 
 
 Separador = vbTab

 Archivo = PathFLog & "Reg " & Format(FechaDesde, "DD-MM-YYYY") & " al " & Format(FechaHasta, "DD-MM-YYYY") & ".txt"

 Set fs = CreateObject("Scripting.FileSystemObject")
 Set freg = fs.CreateTextFile(Archivo, True)


 'StrSql = "Provider=Ifxoledbc;Password=rhpro;Persist Security Info=True;User ID=informix;Data Source=rhpro@rhsco;"
 'OpenConnection StrSql, objConn
 OpenConnection strconexion, objConn
   
  StrSql = " SELECT Gti_Registracion.*,v_empleado.empleg FROM Gti_Registracion " & _
          " INNER JOIN v_empleado ON v_empleado.ternro = gti_registracion.ternro " & _
          " WHERE regfecha >= " & ConvFecha(FechaDesde) & " AND regfecha <= " & ConvFecha(FechaHasta)
 OpenRecordset StrSql, objRs
 
 Do While Not objRs.EOF
    freg.writeline objRs!empleg & Separador & objRs!regfecha & Separador & objRs!reghora
    objRs.MoveNext
 Loop

 freg.Close

End Sub
