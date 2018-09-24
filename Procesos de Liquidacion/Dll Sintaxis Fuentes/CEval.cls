VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CEval - Clase para evaluar expresiones algebraicas

' La Clase implementa un evaluador de expresiones algebraicas.
' Soporta numeros en punto flotante, los operadores estandar (+,-,*,/,^)
' y parentesis para manejar la precedencia de operadores.
' Rudimentariamente soporta el uso de una Tabla de Simbolos definida por el usuario.

' Para usar esta clase en proyectos solo se necesita incluir las clases CEval.cls, CStack.cls
' y CSymbolTable.cls.
' Si no necesita usar la tabla de Simbolos entonces se debe setear la constante
' USE_SYMBOLS #const en False y solamente incluir las clases CEval.cls y CStack.cls en el proyecto.

Option Explicit

'Indica si queremos soportar simbolos (Palabras con un valor particular asociado).
#Const USE_SYMBOLS = True

'Constantes de Estado
Private Const STATE_NONE = 0
Private Const STATE_OPERAND = 1
Private Const STATE_OPERATOR = 2
Private Const STATE_UNARYOP = 3

Private Const UNARY_NEG = "(-)"

Private m_sErrMsg As String

#If USE_SYMBOLS = True Then
    'Objeto Tabla de Simbolos
    Public m_SymbolTable As New CSymbolTable
    Public TablaDeSimbolos As New CSymbolTable
#End If


'Evalua la expresion y retorna el resultado.
Public Function Evaluate(sExpression As String, exito As Boolean, ByVal HacerEvaluacion As Boolean) As Double
    Dim sBuffer As String
    Dim nErrPosition As Integer
    Dim ErrorEnExpresion As Boolean
    
    'Convierte a una expresion en postfijo
    nErrPosition = InfijoAPostfijo(sExpression, sBuffer)

    If nErrPosition Then
        ErrorEnExpresion = True
        'Err.Raise vbObjectError + 1001, , m_sErrMsg & " : Column " & CStr(nErrPosition)
        ErrorPosicion = nErrPosition
        ErrorDescripcion = m_sErrMsg
    Else
        ErrorEnExpresion = False
        ErrorPosicion = 0
        ErrorDescripcion = "Sin Error"
    End If
    
    'Evalua la expresion en postfijo
    If HacerEvaluacion And Not ErrorEnExpresion Then
        Evaluate = DoEvaluate(sBuffer, exito)
    Else
        If Not ErrorEnExpresion Then
            exito = True
        End If
    End If
End Function


Sub EliminarBlancos(Expresion As String, ByRef i As Integer, ch As String, Token As String, Lexema As String, error As String)
Do While Not i > Len(Expresion) And (EsSimboloBlanco(ch))
    ch = Mid$(Expresion, i, 1)
    i = i + 1
Loop

End Sub

Sub RetPalabra(ByVal Expresion As String, ByRef i As Integer, ch As String, Token As String, Lexema As String, error As String)
'Retorna con una palabra leida y con ch = al caracter siguiente en caso de
'ser esto posible, sino con el caracter antes del fin de la expresion.

Do While (Not i > Len(Expresion)) And (EsSimboloDigito(ch) Or EsSimboloChar(ch))
    Lexema = Lexema & ch
    ch = Mid$(Expresion, i, 1)
    i = i + 1
Loop
   
If (EsSimboloDigito(ch) Or EsSimboloLetra(ch) Or EsSimbolo(ch)) Then
    If (EsSimboloDigito(ch) Or EsSimboloLetra(ch)) Then
        Lexema = Lexema & ch
    End If
End If
Token = RetornarPalabraReservada(Token, Lexema)

'retorna blanco si no es pal. reservada
If Token = "" Then
    Token = "Identificador"
End If
'    If Not IdentificadorValido(Lexema) Then
'        Error = "El nombre de parámetro no es válido"
'    End If
'    'token = "Error"
'End If
End Sub

Sub RetNumero(ByVal Expresion As String, i As Integer, ch As String, Token As String, Lexema As String, error As String)
Do While (Not i > Len(Expresion)) And EsSimboloDigito(ch)
    Lexema = Lexema & ch
    ch = Mid$(Expresion, i, 1)
    i = i + 1
Loop
    
If EsSimboloDigito(ch) Then
    Lexema = Lexema & ch
End If
If EsSimboloLetra(ch) Then
    Token = "ERROR"
    error = "Constante numérica no válida"
Else
    Token = "ConstNumerica"
End If
End Sub


Sub MayorOMayorIgual(ByVal Expresion As String, i As Integer, ch As String, Token As String, Lexema As String, error As String)
If Not i > Len(Expresion) Then
    ch = Mid$(Expresion, i, 1)
    i = i + 1
    Select Case ch
    Case "=":
        Lexema = ">="
        Token = "OpMayorIgual"
        If Not i > Len(Expresion) Then
           ch = Mid$(Expresion, i, 1)
           i = i + 1
        End If
    Case Else:
        Lexema = ">"
        Token = "OpMayor"
    End Select
Else
    Lexema = ">"
    Token = "OpMayor"
End If
End Sub

Sub MenorODistintoOMenorIgual(ByVal Expresion As String, i As Integer, ch As String, Token As String, Lexema As String, error As String)
    If Not i > Len(Expresion) Then
        ch = Mid$(Expresion, i, 1)
        i = i + 1
        Select Case ch
        Case ">":
            Lexema = "<>"
            Token = "OpDistinto"
            If Not i > Len(Expresion) Then
                ch = Mid$(Expresion, i, 1)
                i = i + 1
            End If
        Case "="
            Lexema = "<="
            Token = "OpMenorIgual"
            If Not i > Len(Expresion) Then
                ch = Mid$(Expresion, i, 1)
                i = i + 1
            End If
        Case Else:
            Lexema = "<"
            Token = "OpMenor"
        End Select
    Else
        Lexema = "<"
        Token = "OpMenor"
    End If
End Sub
Function RetornarPalabraReservada(ByVal Token As String, ByVal Lexema As String) As String
' Las palabras reservadas son cargadas al inicio en la TS junto con las funciones Predefinidas
' Entonces: si existe tiene asociado el valor correspondiente
'                       (si es parametro ==> el valor del mismo ,
'                       si es una funcion retornara "FUNCION" y
'                       si es una palabra reservada retornará "Reservada")
'           si no retornara "No Definido"
    
    'Ver si la Tabla de Simbolos esta definida
    If eval.m_SymbolTable.IsSymbolDefined(Lexema) Then
            Token = ""
            'Token = UCase(CStr(eval.m_SymbolTable.Value(Lexema)))
            'Lexema = UCase(Lexema)
            
    Else
        Token = "No Definido"
    End If
    RetornarPalabraReservada = Token
End Function


Function NombreTokenSimb(ByVal Expresion As String, i As Integer, ch As String) As String
Select Case ch
    Case "(": NombreTokenSimb = "ParentAbre"
    Case ")": NombreTokenSimb = "ParentCierra"
    Case ">": NombreTokenSimb = "OpMayor"
    Case "<": NombreTokenSimb = "OpMenor"
    Case "=": NombreTokenSimb = "OpIgual"
    Case ",": NombreTokenSimb = "Coma"
    Case "+": NombreTokenSimb = "OpSuma"
    Case "-": NombreTokenSimb = "OpResta"
    Case "*": NombreTokenSimb = "OpProd"
    Case "/": NombreTokenSimb = "OpDiv"
    Case ".": NombreTokenSimb = "Punto"
    Case ";": NombreTokenSimb = "PuntoYComa"
    Case Else:
        NombreTokenSimb = "Desconocido"
End Select
If Not i > Len(Expresion) Then
   ch = Mid$(Expresion, i, 1)
   i = i + 1
End If
End Function

Function IdentificarTipo(ch As String) As String
If EsSimboloLetra(ch) Then
    IdentificarTipo = "Letra"
Else
    If EsSimboloDigito(ch) Then
        IdentificarTipo = "Digito"
    Else
        If EsSimbolo(ch) Then
            IdentificarTipo = "Simbolo"
        Else
            IdentificarTipo = "NoValido"
        End If
    End If
End If
End Function

Sub RecuperarToken(ByVal Expresion As String, ByRef i As Integer, ch As String, Token As String, Lexema As String, error As String)
' Lexema retorna el token leido
' Ch retorna el ultimo caracter leido
' Token retorna el tipo token leido
' I retorna la posicion del ultimo caracter leido
Dim Tipo As String

     Lexema = ""
     Token = ""
     Call EliminarBlancos(Expresion, i, ch, Token, Lexema, error)
 
     
         Tipo = IdentificarTipo(ch)
         Select Case Tipo
         Case "Letra": 'puede ser parametro o un identificador de funcion (identificador) o palabra reservada (AND, OR, etc)
                Call RetPalabra(Expresion, i, ch, Token, Lexema, error)
         Case "Digito":  'es un numero
                Call RetNumero(Expresion, i, ch, Token, Lexema, error)
         Case "Simbolo": 'pueden ser: ( ) < > = + - * / . ,
                Select Case ch
                Case "*", ".", "(", ")", ";", "=", ",", "+", "-", "/":
                    Lexema = ch
                    Token = NombreTokenSimb(Expresion, i, ch)
                    'ch = Mid$(Expresion, i, 1)
                    'i = i + 1
                Case "<":  'puede ser menor o distinto
                        Call MenorODistintoOMenorIgual(Expresion, i, ch, Token, Lexema, error)
                Case ">": 'Puede ser mayor o mayor igual
                        Call MayorOMayorIgual(Expresion, i, ch, Token, Lexema, error)
                Case Else:
                    Token = "ERROR"
                    error = "Caracter no válido en expresión"
                End Select 'CASE CH
        Case Else:
            Token = "ERROR"
            error = "Caracter no válido en expresión"
        End Select 'CASE TIPO}
    
    If i > Len(Expresion) Then
        i = i + 1
    End If
End Sub

'Convierte una expresion en InFijo a una expresion en PostFijo de tokens separados por espacio.
Private Function InfijoAPostfijo(sExpression As String, sBuffer As String) As Integer
    Dim i As Integer, ch As String, sTemp As String
    Dim nCurrState As Integer, nParenCount As Integer
    Dim bDecPoint As Boolean
    Dim stkTokens As New CStack
    Dim Exp2 As String
    Dim exito As Boolean
    Dim Token As String
    Dim error As String
    Dim Lexema As String
    
    sExpression = sExpression & " "
    error = ""
    nCurrState = STATE_NONE
    nParenCount = 0
    i = 1
    ch = Mid$(sExpression, i, 1)
    i = i + 1
    'Call RecuperarToken(sExpression, i, ch, Token, Lexema, Error)
    
    Do Until i > Len(sExpression) + 1
        Call RecuperarToken(sExpression, i, ch, Token, Lexema, error)
        'If i > Len(sExpression) + 1 Then
        '    i = i + 1
        'End If
        Select Case Token
        Case "ParentAbre":
            'no puede seguir un operando
            If nCurrState = STATE_OPERAND Then
                m_sErrMsg = "Se esperaba Operador"
                GoTo EvalError
            End If
            'permite operador unario Menos(-)antes "("
            If nCurrState = STATE_UNARYOP Then
                nCurrState = STATE_OPERATOR
            End If
            'Apila Parentesis en la pila
            stkTokens.Push Lexema
            'Lleva la cuenta de los parentesis en la Pila
            nParenCount = nParenCount + 1
        
        Case "ParentCierra":
            'debe seguir a un operando
            If nCurrState <> STATE_OPERAND Then
                m_sErrMsg = "se esperaba Operando"
                GoTo EvalError
            End If
            'Deben coincidir los parentesis.
            If nParenCount = 0 Then
                m_sErrMsg = "Parentesis de cierre sin parentesis de apertura"
                GoTo EvalError
            End If
            'desapila todos los operadores hasta encontrar "(".
            sTemp = stkTokens.Pop
            Do Until sTemp = "("
                sBuffer = sBuffer & sTemp & " "
                sTemp = stkTokens.Pop
            Loop
            'Lleva la cuenta de los parentesis en la pila.
            nParenCount = nParenCount - 1
        
        Case "OpMas", "OpResta", "OpProd", "OpDiv", "OpIgual", "OpMayorIgual", "OpMayor", "OpMenor", "OpMenorIgual", "OpDistinto", "AND", "OR":
            'necesito un bit extra para maejar el menos unario
            If nCurrState = STATE_OPERAND Then
                'Desapila operadores con precedencia >= al operador en lexema
                Do While stkTokens.StackSize > 0
                    If GetPrecedence(stkTokens.GetPopValue) < GetPrecedence(Lexema) Then
                        Exit Do
                    End If
                    sBuffer = sBuffer & stkTokens.Pop & " "
                Loop
                'Apila l nuevo operando
                stkTokens.Push Lexema
                nCurrState = STATE_OPERATOR
            ElseIf nCurrState = STATE_UNARYOP Then
                'No se permiten dos operadores unarios seguidos
                m_sErrMsg = "Se esperaba Operando"
                GoTo EvalError
            Else
                'Test para el Operador Unario
                If Token = "OpResta" Then
                    'Apila el menos unario
                    stkTokens.Push UNARY_NEG
                    nCurrState = STATE_UNARYOP
                ElseIf Token = "OpMas" Then
                    'Simplemente se ignora el operador unario +
                    nCurrState = STATE_UNARYOP
                Else
                    m_sErrMsg = "Se esperaba Operando"
                    GoTo EvalError
                End If
            End If
        
        Case "AND", "OR":
            If nCurrState = STATE_OPERAND Then
                'Desapila operadores con precedencia >= al operador en lexema
                Do While stkTokens.StackSize > 0
                    sBuffer = sBuffer & stkTokens.Pop & " "
                Loop
                'Apila l nuevo operando
                stkTokens.Push Lexema
                nCurrState = STATE_OPERATOR
            Else
                'No se permiten dos operadores unarios seguidos
                m_sErrMsg = "Se esperaba Operando"
                GoTo EvalError
            End If
        
        Case "ConstNumerica":
            'No puede seguir otro operando
            If nCurrState = STATE_OPERAND Then
                m_sErrMsg = "Se esperaba Operador"
                GoTo EvalError
            End If
            sBuffer = sBuffer & Lexema & " "
            nCurrState = STATE_OPERAND
        
        Case "Identificador":
                'Un Simbol no puede seguir a otro operando
                If nCurrState = STATE_OPERAND Then
                    m_sErrMsg = "Se esperaba Operador"
                    GoTo EvalError
                End If
                ' reviso que el parametro exista en la TS
                
                If Lexema <> "No Definido" Then
                    If Lexema = "SI" Then
                        'Es una funcion ==> la evaluo
                            Call EvaluarFuncion(sExpression, ch, i, Lexema, Exp2, exito)
                        If Not exito Then
                            GoTo EvalError
                        End If
    
                        sBuffer = sBuffer & Exp2 & " "
                        nCurrState = STATE_OPERAND
                    Else
                        Lexema = UCase(CStr(eval.m_SymbolTable.Value(Lexema)))
                        sBuffer = sBuffer & Lexema & " "
                        nCurrState = STATE_OPERAND
                    End If
                Else
                    m_sErrMsg = "Simbolo no definido : '" & sTemp & "'"
                    GoTo EvalError
                End If
        Case Else
        
        End Select
    Loop
    
    'La expresion no puede terminar con un operador
    If nCurrState = STATE_OPERATOR Or nCurrState = STATE_UNARYOP Then
        m_sErrMsg = "Se esperaba Operando"
        GoTo EvalError
    End If
    'Chequeo que los parentesis esten balanceados
    If nParenCount > 0 Then
        m_sErrMsg = "Se esperaba parentesis que cierra"
        GoTo EvalError
    End If
    'Retorna los operadores restantes en la pila
    Do Until stkTokens.StackSize = 0
        sBuffer = sBuffer & stkTokens.Pop & " "
    Loop
    'Marca que no hay error
    InfijoAPostfijo = 0
    Exit Function
    
EvalError:
    'Reporta la posicion del Error
    InfijoAPostfijo = i
    Exit Function
End Function


Public Sub EvaluarFuncion(ByRef Expresion As String, ByRef ch As String, ByRef i As Integer, ByVal nombre As String, ByRef exp As String, ByRef exito As Boolean)
'Dim ch As String
Dim ContadorParentesis As Integer
Dim sIF As String
Dim Continua As Boolean

ContadorParentesis = 0
sIF = ""
Select Case UCase(nombre)
Case "SI":
    ' Buscar Condicion
    'ch = Mid$(Expresion, i, 1)
    'i = i + 1
    If EsSimboloParAbre(ch) Then
        sIF = ch
        ContadorParentesis = 1
        Continua = True
        ch = Mid$(Expresion, i, 1)
        i = i + 1
    Else
        Continua = False
    End If

    Do Until (EsSimboloParCierra(ch) And ContadorParentesis = 0 And Continua) Or (i > Len(Expresion))
        sIF = sIF & ch
        If i > Len(Expresion) Then Exit Do
        ch = Mid$(Expresion, i, 1)
        i = i + 1
        If EsSimboloParAbre(ch) Then
            ContadorParentesis = ContadorParentesis + 1
        End If
        If EsSimboloParCierra(ch) Then
            ContadorParentesis = ContadorParentesis - 1
        End If
        
    Loop
    If Not EsSimboloBlanco(ch) Then
        sIF = sIF & ch
        ch = Mid$(Expresion, i, 1)
        i = i + 1
    End If
    
    Call EvaluarIf(sIF, exp, exito)

Case Else

End Select
End Sub

Public Sub EvaluarIf(ByVal sIF As String, ByRef exp As String, ByRef exito As Boolean)
' Al momento de entrar sif tiene todo el interior de la expresion condicional, los tres paramteros del SI(Cond,Exp1,Exp2)
Dim ch As String
Dim pos As Integer
Dim SCondicion As String
Dim SExp1 As String
Dim SExp2 As String
Dim ContadorParentesis As Integer
Dim Posfinal As Integer
Dim CantComas As Integer
Dim tok As String
Dim lex As String
Dim IFerror As String

    exito = False
    CantComas = 0
    
    ' le tengo que sacar el primer y ultimo parentesis
    Posfinal = Len(sIF)
    ch = Mid$(sIF, Posfinal, 1)
    Do Until EsSimboloParCierra(ch)
        Posfinal = Posfinal - 1
        ch = Mid$(sIF, Posfinal, 1)
    Loop
    sIF = Mid$(sIF, 2, (Len(sIF) - (Len(sIF) - Posfinal)) - 2)

    pos = 1
    ch = Mid$(sIF, pos, 1)
    pos = pos + 1
    Do While Not EsSimboloPuntoyComa(ch)
        SCondicion = SCondicion & ch
        ch = Mid$(sIF, pos, 1)
        pos = pos + 1
        If pos > Len(sIF) Then Exit Do
    Loop

    'salteo el punto y coma
    ch = Mid$(sIF, pos, 1)
    pos = pos + 1

    Call RecuperarToken(sIF, pos, ch, tok, lex, IFerror)
    
    ' Si leo alguna funcion tengo que saltear tantas comas como parametros tenga esa funcion
    If UCase(tok) = UCase("identificador") Then
        Select Case UCase(lex)
        Case "SI":
            CantComas = CantComas + 2
        Case Else:
            CantComas = CantComas + 1
        End Select
    End If
    
    Do Until (lex = ";" And CantComas = 0) Or (pos > Len(sIF)) Or (IFerror <> "" And IFerror <> "SinError")
        SExp1 = SExp1 & lex
        If lex = ";" Then
            CantComas = CantComas - 1
        End If
'        ch = Mid$(sIF, pos, 1)
'        pos = pos + 1
        Call RecuperarToken(sIF, pos, ch, tok, lex, IFerror)
        'If pos > Len(sIF) Then Exit Do
    Loop

    'salteo la coma
    'ch = Mid$(sIF, pos, 1)
    'pos = pos + 1

    If pos > Len(sIF) And Not EsSimboloBlanco(ch) Then
        SExp2 = SExp2 & ch
    End If
    'ContadorParentesis = 0
    Do While Not pos > Len(sIF) And (IFerror = "" Or IFerror = "SinError")
        SExp2 = SExp2 & ch
        ch = Mid$(sIF, pos, 1)
        pos = pos + 1
        If pos > Len(sIF) Then
            SExp2 = SExp2 & ch
        End If
    Loop

If IFerror = "" Or IFerror = "SinError" Then
    If EvaluarCondicion(SCondicion, exito) Then
        If exito Then
            exp = CSng(eval.Evaluate(SExp1, exito, True))
        End If
    Else
        If exito Then
            exp = CSng(eval.Evaluate(SExp2, exito, True))
        End If
    End If
Else
    exito = False
End If
End Sub

Public Function EvaluarCondicion(ByVal SCondicion As String, ByRef exito As Boolean) As String
Dim ResLogico As Boolean
' <Condicion> :: = Exp OpRel Exp CondLog
' <OpRel>::= "<" | ">" | "=" | "<=" | ">=" | "<>"
' <CondLog>::= Vacio | OpLog Condicion
' <OpLog>::= "AND" | "OR"

ResLogico = CBool(eval.Evaluate(SCondicion, exito, True))

If exito Then
    EvaluarCondicion = ResLogico
Else
    EvaluarCondicion = False
End If
End Function

'Retorna un numero que indica la precedencia relativa del operador.
Private Function GetPrecedence(ch As String) As Integer
    Select Case ch
        Case "+", "-"
            GetPrecedence = 1
        Case "*", "/"
            GetPrecedence = 2
        Case "^"
            GetPrecedence = 3
        Case UNARY_NEG
            GetPrecedence = 10
        Case Else
            GetPrecedence = 0
    End Select
End Function

'Evalua la expresion y retorna el resultado.
'Se asume que la expresion ha sido convertida a PostFijo.

Private Function DoEvaluate(sExpression As String, exito As Boolean) As Double
    Dim i As Integer, j As Integer, stkTokens As New CStack
    Dim sTemp As String, Op1 As Variant, Op2 As Variant
    Dim R
    'Localiza el primer token
    i = 1
    j = InStr(sExpression, " ")
    Do Until j = 0
        'Extrae un token de la expresion
        sTemp = Mid$(sExpression, i, j - i)
        If IsNumeric(sTemp) Then
            'si es operando, lo mete en la Pila
            stkTokens.Push (sTemp)
        Else
            'Si es operador, realiza el calculo
            Select Case sTemp
                Case "+"
                    stkTokens.Push stkTokens.Pop + stkTokens.Pop
                Case "-"
                    Op1 = stkTokens.Pop
                    Op2 = stkTokens.Pop
                    stkTokens.Push Op2 - Op1
                Case "*"
                    stkTokens.Push stkTokens.Pop * stkTokens.Pop
                Case "/"
                    Op1 = stkTokens.Pop
                    Op2 = stkTokens.Pop
                    stkTokens.Push Op2 / Op1
                Case ">"
                    Op1 = stkTokens.Pop
                    Op2 = stkTokens.Pop
                    stkTokens.Push CBool(Op2 > Op1)
                Case "<"
                    Op1 = stkTokens.Pop
                    Op2 = stkTokens.Pop
                    stkTokens.Push CBool(Op2 < Op1)
                Case "="
                    Op1 = stkTokens.Pop
                    Op2 = stkTokens.Pop
                    stkTokens.Push CBool(Op2 = Op1)
                Case ">="
                    Op1 = stkTokens.Pop
                    Op2 = stkTokens.Pop
                    stkTokens.Push CBool(Op2 >= Op1)
                Case "<="
                    Op1 = stkTokens.Pop
                    Op2 = stkTokens.Pop
                    stkTokens.Push CBool(Op2 <= Op1)
                Case "<>"
                    Op1 = stkTokens.Pop
                    Op2 = stkTokens.Pop
                    stkTokens.Push CBool(Op2 <> Op1)
                Case "AND"
                    Op1 = stkTokens.Pop
                    Op2 = stkTokens.Pop
                    stkTokens.Push CBool(Op2 And Op1)
                Case "OR"
                    Op1 = stkTokens.Pop
                    Op2 = stkTokens.Pop
                    stkTokens.Push CBool(Op2 Or Op1)
                Case UNARY_NEG
                    stkTokens.Push -stkTokens.Pop
                Case Else
                    'No deberia ocurrir (tokens invalidos en la expresion transformadsa por funcion InfixToPostfix)
                    ErrorEnExpresion = True
                    'Err.Raise vbObjectError + 1002, , "Token incorrecto en la evaluación: " & sTemp
            End Select
        End If
        'Siguiente token
        i = j + 1
        j = InStr(i, sExpression, " ")
    Loop
    'El item que quedó en la pila contiene el resultado
    If stkTokens.StackSize > 0 Then
        'R = (stkTokens.Pop)
        DoEvaluate = (stkTokens.Pop)
        exito = True
    Else
        'Expresión; retorna 0
        DoEvaluate = 0
        exito = False
    End If
End Function

'Retorna True o False segun si el sChar es valido para ser usado como
' primer caracter de nombre en la tabla de simbolos
Private Function IsSymbolCharFirst(sChar As String) As Boolean
    Dim c As String

    c = UCase(Left(sChar, 1))
    If (c >= "A" And c <= "Z") Or InStr("_", c) Then
        IsSymbolCharFirst = True
    Else
        IsSymbolCharFirst = False
    End If
End Function

'Retorna True o False segun si el sChar es valido para ser usado como
'caracter de nombre en la tabla de simbolos
'Private Function IsSymbolChar(sChar As String) As Boolean
'    Dim c As String
'
'    c = UCase(Left(sChar, 1))
'    If (c >= "A" And c <= "Z") Or InStr("0123456789_", c) Then
'        IsSymbolChar = True
'    Else
'        IsSymbolChar = False
'    End If
'End Function

'Retorna True o False segun si el sChar es valido para ser usado como
'caracter de nombre en la tabla de simbolos
Private Function EsSimboloChar(sChar As String) As Boolean
    Dim c As String

    c = UCase(Left(sChar, 1))
    If (c >= "A" And c <= "Z") Or InStr("0123456789_", c) Then
        EsSimboloChar = True
    Else
        EsSimboloChar = False
    End If
End Function

Private Function EsSimboloLetra(sChar As String) As Boolean
    Dim c As String

    c = UCase(Left(sChar, 1))
    If (c >= "A" And c <= "Z") Then
        EsSimboloLetra = True
    Else
        EsSimboloLetra = False
    End If
End Function

Private Function EsSimboloDigito(ByVal ch As String) As Boolean
    
    If InStr("0123456789,.", ch) Then
        EsSimboloDigito = True
    Else
        EsSimboloDigito = False
    End If
End Function

Private Function EsSimboloBlanco(ByVal ch As String) As Boolean
    
Select Case Asc(ch)
Case 8: ' backspace
    EsSimboloBlanco = True
Case 9: ' Tab
    EsSimboloBlanco = True
Case 10: 'Nueva Linea
    EsSimboloBlanco = True
Case 12: 'Form Feed
    EsSimboloBlanco = True
Case 13: 'Retorno de Carro
    EsSimboloBlanco = True
Case 32: 'Espacio en Blanco
    EsSimboloBlanco = True
Case Else:
    EsSimboloBlanco = False
End Select
    
End Function

Function EsSimbolo(ch As String) As Boolean
Dim SimbolosValidos As String
Dim Posicion

    SimbolosValidos = "()<>=,+-*/.;"
    
    Posicion = InStr(1, SimbolosValidos, ch)
    
    EsSimbolo = IIf((Posicion = 0) Or IsNull(Posicion), False, True)
End Function

'Private Function IsSymbolClosePar(sChar As String) As Boolean
'    Dim c As String
'
'    c = UCase(Left(sChar, 1))
'    If c = ")" Then
'        IsSymbolClosePar = True
'    Else
'        IsSymbolClosePar = False
'    End If
'End Function

Private Function EsSimboloParCierra(sChar As String) As Boolean
    Dim c As String

    c = UCase(Left(sChar, 1))
    If c = ")" Then
        EsSimboloParCierra = True
    Else
        EsSimboloParCierra = False
    End If
End Function

Private Function EsSimboloComa(sChar As String) As Boolean
    Dim c As String

    c = UCase(Left(sChar, 1))
    If c = "," Then
        EsSimboloComa = True
    Else
        EsSimboloComa = False
    End If
End Function

Private Function EsSimboloPuntoyComa(sChar As String) As Boolean
    Dim c As String

    c = UCase(Left(sChar, 1))
    If c = ";" Then
        EsSimboloPuntoyComa = True
    Else
        EsSimboloPuntoyComa = False
    End If
End Function

Private Function EsSimboloOpRelacional(sChar As String) As Boolean
    Dim c As String
    
    EsSimboloOpRelacional = False
    
    c = UCase(Left(sChar, 1))
    If c = "=" Then
        EsSimboloOpRelacional = True
    End If
    If c = "<" Then
        ' tengo que leer uno mas
        EsSimboloOpRelacional = True
    End If
    If c = ">" Then
        ' tengo que leer uno mas
        EsSimboloOpRelacional = True
    End If

End Function


'Private Function IsSymbolOpenPar(sChar As String) As Boolean
'    Dim c As String
'
'    c = UCase(Left(sChar, 1))
'    If c = "(" Then
'        IsSymbolOpenPar = True
'    Else
'        IsSymbolOpenPar = False
'    End If
'End Function

Private Function EsSimboloParAbre(sChar As String) As Boolean
    Dim c As String

    c = UCase(Left(sChar, 1))
    If c = "(" Then
        EsSimboloParAbre = True
    Else
        EsSimboloParAbre = False
    End If
End Function

