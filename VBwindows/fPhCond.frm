VERSION 5.00
Begin VB.Form fPhCond 
   Caption         =   "Multipar"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2700
   Icon            =   "fPhCond.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   2700
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bAvanti 
      Caption         =   "&Avanti >"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton bIndietro 
      Caption         =   "< &Indietro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.OptionButton oConducibilita 
      Caption         =   "&Conducibilit�"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.OptionButton oPh 
      Caption         =   "&pH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Scegliere un sensore fra pH e Conducibilit�."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   255
      TabIndex        =   4
      Top             =   0
      Width           =   2250
   End
End
Attribute VB_Name = "fPhCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    If Canale(3).Attivo = True Then
        oPh.Value = True
        Canale(3).Attivo = True
        Canale(1).Attivo = False
    Else
        oConducibilita.Value = True
        Canale(1).Attivo = True
        Canale(3).Attivo = False
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Canale(1).Attivo = True
    Canale(3).Attivo = True
    Unload Me
    fMain.Show
End Sub

Private Sub bAvanti_Click()
    Label1.Caption = "Attendere..."
    bAvanti.Enabled = False
    bIndietro.Enabled = False
    
    Load fVislChnl

    DoEvents
    Me.Hide
    Unload Me
    fVislChnl.Show
End Sub

Private Sub bIndietro_Click()
    Unload Me
    frmOptions2.Show
End Sub

Private Sub oConducibilita_Click()
    Canale(1).Attivo = True
    Canale(3).Attivo = False
End Sub

Private Sub oPh_Click()
    Canale(1).Attivo = False
    Canale(3).Attivo = True
End Sub
