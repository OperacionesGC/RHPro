VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9465
   LinkTopic       =   "Form2"
   ScaleHeight     =   6630
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Funciones Validas "
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      Begin VB.OptionButton Option4 
         Caption         =   "Valor Absoluto (ABS)"
         Height          =   375
         Left            =   6720
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Truncado (TRUNC)"
         Height          =   375
         Left            =   4560
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Redondeo (RED)"
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Condicional (SI)"
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.OLE OLE1 
      AutoActivate    =   1  'GetFocus
      Class           =   "Word.Document.8"
      Height          =   5655
      Left            =   120
      OleObjectBlob   =   "Formulario_ayuda.frx":0000
      SourceDoc       =   "C:\Visual\Procesos de Liquidacion\Dll sintaxis\Funcion_SI.doc"
      TabIndex        =   1
      Top             =   840
      Width           =   9255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'    If Option1.Value Then   'Funcion SI
'        OLE1.SourceDoc = "C:\Visual\Procesos de Liquidacion\Dll sintaxis\Funcion_SI.doc"
'        OLE1.Refresh
'    End If
'    If Option2.Value Then   'Funcion Redondeo
'        OLE1.SourceDoc = "C:\Visual\Procesos de Liquidacion\Dll sintaxis\Funcion_Red.doc"
'        OLE1.Refresh
'    End If
'    If Option3.Value Then   'Funcion Truncado
'        OLE1.SourceDoc = "C:\Visual\Procesos de Liquidacion\Dll sintaxis\Funcion_Trunc.doc"
'        OLE1.Refresh
'    End If
'    If Option4.Value Then   'Funcion ABS
'        OLE1.SourceDoc = "C:\Visual\Procesos de Liquidacion\Dll sintaxis\Funcion_ABS.doc"
'        OLE1.Refresh
'    End If
End Sub

Private Sub Option1_Click()
    If Option1.Value Then   'Funcion SI
        OLE1.SourceDoc = "C:\Visual\Procesos de Liquidacion\Dll sintaxis\Funcion_SI.doc"
        OLE1.Refresh
        OLE1.SetFocus
    End If
End Sub

Private Sub Option2_Click()
    If Option2.Value Then   'Funcion Redondeo
        OLE1.SourceDoc = "C:\Visual\Procesos de Liquidacion\Dll sintaxis\Funcion_Red.doc"
        OLE1.Refresh
        OLE1.SetFocus
    End If
End Sub

Private Sub Option3_Click()
    If Option3.Value Then   'Funcion Truncado
        OLE1.SourceDoc = "C:\Visual\Procesos de Liquidacion\Dll sintaxis\Funcion_Trunc.doc"
        OLE1.Refresh
        OLE1.SetFocus
    End If
End Sub

Private Sub Option4_Click()
    If Option4.Value Then   'Funcion ABS
        OLE1.SourceDoc = "C:\Visual\Procesos de Liquidacion\Dll sintaxis\Funcion_ABS.doc"
        OLE1.Refresh
        OLE1.SetFocus
    End If
End Sub
