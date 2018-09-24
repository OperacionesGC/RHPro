VERSION 5.00
Begin VB.Form FormEncript 
   BackColor       =   &H80000013&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RHProX2 - Encript"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7665
   FillColor       =   &H80000013&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   120
      Picture         =   "FormEncript.frx":0000
      ScaleHeight     =   1035
      ScaleWidth      =   2595
      TabIndex        =   12
      Top             =   0
      Width           =   2655
   End
   Begin VB.TextBox StrDestino 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2520
      Width           =   7335
   End
   Begin VB.TextBox strOrigen 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "FormEncript.frx":146B
      Top             =   1440
      Width           =   7335
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   3480
      Width           =   7335
      Begin VB.CommandButton BtnDecript 
         Caption         =   "Desencriptar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.Frame FrameSeed 
         Caption         =   "Semilla de Encriptacion "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4560
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   2535
         Begin VB.CommandButton btnok 
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   10
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox strSeed 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Text            =   "56238"
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CommandButton btnConf 
         Caption         =   "Configurar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton BtnSalir 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton BtnEncript 
         Caption         =   "Encriptar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label2 
      Caption         =   "String resultante"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "String de origen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "FormEncript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnConf_Click()
    FrameSeed.Visible = True
End Sub

Private Sub BtnDecript_Click()

    On Error GoTo ME_Local
    StrDestino = Decrypt(strSeed, strOrigen)
    Exit Sub
ME_Local:
    MsgBox "No se puede desencriptar", vbCritical
End Sub

Private Sub BtnEncript_Click()
    On Error GoTo ME_Local
    StrDestino = Encrypt(strSeed, strOrigen)
    Exit Sub
    
ME_Local:
    MsgBox "No se puede encriptar", vbCritical
End Sub

Private Sub btnok_Click()
    FrameSeed.Visible = False
End Sub

Private Sub BtnSalir_Click()
End
End Sub

Private Sub Form_Load()
    FrameSeed.Visible = False
End Sub
