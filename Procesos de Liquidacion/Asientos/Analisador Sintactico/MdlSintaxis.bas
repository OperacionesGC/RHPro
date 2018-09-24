Attribute VB_Name = "AnalisadorSintactico"
Option Explicit

Global eval As New Sintaxis.AnalisadorSintactico


Global ErrorPosicion As Integer
Global ErrorDescripcion As String
Global StrSql As String
Global NroFormula As Long
Global ErrorEnExpresion As Boolean


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: procedimiento inicial del Evaluador de expresiones del liquidador.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim exito As Boolean
Dim evaluar As Boolean
Dim Expresion As String
Dim rs_for_chequeo As New ADODB.Recordset
Dim Resultado As Single

    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    
    'Abro la conexion
    OpenConnection strconexion, objConn
    
    strCmdLine = Command()
    If IsNumeric(strCmdLine) Then
        NroFormula = strCmdLine
    Else
        Exit Sub
    End If
    
    MsgBox "Empieza con el chequeo de la formula : " & NroFormula, vbCritical
    
    'Cambio el estado de la formula a chequeando
    
    StrSql = "UPDATE for_chequeo SET forEstado = 0 , forErrPos = 0, forErrDesc = ' ' WHERE fornro = " & NroFormula
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = " SELECT * FROM for_chequeo " & _
             " WHERE for_chequeo.fornro = " & NroFormula
    OpenRecordset StrSql, rs_for_chequeo
    
    If Not rs_for_chequeo.EOF Then
        
        Call CargarTablaParametros
        Expresion = rs_for_chequeo!forexpresion
    
        Resultado = CSng(eval.Evaluate(Trim(Expresion), exito, CBool(rs_for_chequeo!forresultado)))
    End If
    
    'Cambio el estado de la formula a Termino
    StrSql = "UPDATE for_chequeo SET forEstado = 1 , forErrPos = " & ErrorPosicion & ", forErrDesc = '" & ErrorDescripcion & "', forValor = " & Resultado & " WHERE fornro = " & NroFormula
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub


Public Sub CargarTablaParametros()
' carga la tabla de simbolos con los parametros en wf_tpa

Dim symbols As New CSymbolTable
Dim rs_wf_tpa As New ADODB.Recordset
Dim NParametro As String
Dim valor As Single
Dim rs_FunFormulas As New ADODB.Recordset
Dim rs_For_Tpa As New ADODB.Recordset

    Set eval.m_SymbolTable = symbols
     
    'Aca deberia cargar todas las funciones validas
    StrSql = "SELECT * FROM funformula "
    OpenRecordset StrSql, rs_FunFormulas

    Do While Not rs_FunFormulas.EOF
        'por ahora las inserto a dedo
        'eval.m_SymbolTable.Add "SI", "Funcion"
        
        eval.m_SymbolTable.Add UCase(rs_FunFormulas!fundesabr), "Funcion"
        
        rs_FunFormulas.MoveNext
    Loop
    
  
    ' Cada parametro en wf_tpa se inserta como un simbolo en la tabla
    ' Resolucion de los parámetros de la fórmula
    StrSql = "SELECT * FROM for_tpa " & _
             " WHERE for_tpa.fornro = " & NroFormula & _
             " ORDER BY for_tpa.ftorden"
    OpenRecordset StrSql, rs_For_Tpa
    
    Do While Not rs_For_Tpa.EOF
        NParametro = Trim("par" & Format(rs_For_Tpa!tpanro, "00000"))
        'valor = rs_For_Tpa!valor
        valor = 1
        
        eval.m_SymbolTable.Add NParametro, valor
        
        rs_For_Tpa.MoveNext
    Loop
    
    
End Sub
