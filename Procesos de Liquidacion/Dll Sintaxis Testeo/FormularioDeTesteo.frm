VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9435
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEvaluate 
      Caption         =   "E&valuar"
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Frame FrameOpciones 
      Caption         =   "Opciones de Chequeo "
      Height          =   1455
      Left            =   6720
      TabIndex        =   5
      Top             =   360
      Width           =   2655
      Begin VB.CheckBox CheckEvaluar 
         Caption         =   "Evaluar Expresion"
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   480
         Value           =   1  'Checked
         Width           =   1695
      End
   End
   Begin VB.TextBox txtExpression 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   0
      TabIndex        =   2
      Text            =   "((45 * 7) + 14) * 12"
      Top             =   360
      Width           =   6495
   End
   Begin VB.TextBox txtResult 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   525
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0"
      Top             =   2280
      Width           =   6495
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "S&alir"
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Expresión"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Resultado"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   1920
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEvaluate_Click()
Dim exito As Boolean
Dim Resultado As String
Dim HacerEvaluacion As Boolean

    Set eval = New AnalisadorSintactico

    HacerEvaluacion = CBool(CheckEvaluar.Value)
    Call CargarTablaParametros
    Resultado = eval.Evaluate(txtExpression, exito, HacerEvaluacion)
    If exito Then
        If HacerEvaluacion Then
            txtResult = Resultado
        Else
            txtResult = "Sintaxis Correcta"
        End If
    Else
        txtResult = "Sintaxis Incorrecta : " & eval.ErrMsg & " en posicion : " & eval.ErrPosicion
    End If
End Sub

Public Sub CargarTablaParametros()
' carga la tabla de simbolos con los parametros en wf_tpa

Dim symbols As New CSymbolTable

    Set eval.m_SymbolTable = symbols
     
    eval.m_SymbolTable.Add "SI", "Funcion"
    eval.m_SymbolTable.Add "ABS", "Funcion"
    eval.m_SymbolTable.Add "RED", "Funcion"
    eval.m_SymbolTable.Add "TRUNC", "Funcion"
    eval.m_SymbolTable.Add "AND", "OP LOGICO"
    eval.m_SymbolTable.Add "OR", "OP LOGICO"
    eval.m_SymbolTable.Add "Minimo", "0.12251"
    eval.m_SymbolTable.Add "Maximo", "2"
    
End Sub


Private Sub cmdExit_Click()
    End
End Sub

