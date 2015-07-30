VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmOptions2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opzioni Test"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   5550
   ClientWidth     =   6195
   Icon            =   "frmOptions2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Decimali"
      Height          =   615
      Left            =   2280
      TabIndex        =   29
      Top             =   4320
      Width           =   855
      Begin VB.TextBox tDecimali 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Text            =   "3"
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frequenza aggiornamento"
      Height          =   615
      Left            =   120
      TabIndex        =   27
      Top             =   4320
      Width           =   2055
      Begin VB.TextBox tFreqAgg 
         Height          =   285
         Left            =   720
         TabIndex        =   1
         Text            =   "1"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "secondi"
         Height          =   255
         Left            =   1080
         TabIndex        =   28
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Applica"
      Height          =   375
      Left            =   4560
      TabIndex        =   19
      ToolTipText     =   "Conferma i dati immessi"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton bIndietro 
      Caption         =   "< &Indietro"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      ToolTipText     =   "Conferma  ed esce"
      Top             =   4440
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Canale 0"
      Height          =   3375
      Left            =   350
      TabIndex        =   16
      Top             =   700
      Width           =   5415
      Begin VB.Frame Frame3 
         Caption         =   "Setup Grandezze fisiche"
         Height          =   1425
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   3855
         Begin VB.TextBox tValMin 
            Height          =   285
            Left            =   2160
            TabIndex        =   10
            Text            =   "0"
            Top             =   210
            Width           =   1575
         End
         Begin VB.TextBox tValMax 
            Height          =   285
            Left            =   2160
            TabIndex        =   11
            Text            =   "0"
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox tValoff 
            Height          =   285
            Left            =   2160
            TabIndex        =   12
            Text            =   "0"
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Valore minimo"
            Height          =   252
            Left            =   120
            TabIndex        =   26
            Top             =   300
            Width           =   1932
         End
         Begin VB.Label Label6 
            Caption         =   "Valore massimo"
            Height          =   228
            Left            =   120
            TabIndex        =   25
            Top             =   624
            Width           =   1932
         End
         Begin VB.Label Label7 
            Caption         =   "Offset"
            Height          =   252
            Left            =   120
            TabIndex        =   24
            Top             =   984
            Width           =   1932
         End
      End
      Begin VB.CommandButton bSalva 
         Caption         =   "&Salva setup"
         Height          =   615
         Left            =   4320
         TabIndex        =   14
         Top             =   960
         Width           =   852
      End
      Begin VB.CommandButton bLeggi 
         Caption         =   "Ca&rica setup"
         Height          =   615
         Left            =   4320
         TabIndex        =   13
         Top             =   1710
         Width           =   855
      End
      Begin VB.TextBox tUnita 
         Height          =   285
         Left            =   3000
         MaxLength       =   4
         TabIndex        =   7
         ToolTipText     =   "Qui va messa l'unit� di misura (es. mm)"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox tNome 
         Height          =   285
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   6
         ToolTipText     =   "Nome canale"
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox oAttivo 
         Caption         =   "Attivo"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.Frame Frame2 
         Caption         =   "Setup ADC"
         Height          =   945
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   3855
         Begin VB.TextBox tBitMin 
            Height          =   285
            Left            =   2160
            TabIndex        =   8
            Text            =   "0"
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox tBitMax 
            Height          =   285
            Left            =   2160
            TabIndex        =   9
            Text            =   "4095"
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Valore bit min (ADC)"
            Height          =   252
            Left            =   144
            TabIndex        =   23
            Top             =   288
            Width           =   1452
         End
         Begin VB.Label Label4 
            Caption         =   "Valore bit max (ADC)"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Unit� di misura"
         Height          =   255
         Left            =   3000
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000A&
         Caption         =   "Nome"
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "A&nnulla"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Continua >"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      ToolTipText     =   "Conferma  ed esce"
      Top             =   4440
      Width           =   975
   End
   Begin ComctlLib.TabStrip tbsOptions 
      Height          =   4245
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7488
      MultiRow        =   -1  'True
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   18
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "1"
            Key             =   "chan0"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Imposta le opzioni per il Canale 1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "2"
            Key             =   "Chan1"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Imposta le opzioni per il Canale 2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "3"
            Key             =   "Chan2"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Imposta le opzioni per il Canale 3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "4"
            Key             =   "Chan3"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Imposta le opzioni per il Canale 4"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "5"
            Key             =   "Chan4"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Imposta le opzioni per il Canale 5"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "6"
            Key             =   "Chan5"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Imposta le opzioni per il Canale 6"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "7"
            Key             =   "Chan6"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Imposta le opzioni per il Canale 7"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "8"
            Key             =   "Chan7"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Imposta le opzioni per il Canale 8"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "9"
            Key             =   "Chan8"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Imposta le opzioni per il Canale 9"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "10"
            Key             =   "Chan9"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Imposta le opzioni per il Canale 10"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "11"
            Key             =   "Chan10"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Imposta le opzioni per il Canale 11"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab12 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "12"
            Key             =   "Chan11"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Imposta le opzioni per il Canale 12"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab13 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "13"
            Key             =   "Chan12"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Imposta le opzioni per il Canale 13"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab14 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "14"
            Key             =   "Chan13"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Imposta le opzioni per il Canale 14"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab15 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "15"
            Key             =   "Chan14"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Imposta le opzioni per il Canale 15"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab16 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "16"
            Key             =   "Chan15"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Imposta le opzioni per il Canale 16"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab17 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "17"
            Key             =   "Chan16"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Imposta le opzioni per il Canale 17"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab18 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "18"
            Key             =   "Chan17"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Imposta le opzioni per il Canale 18"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOptions2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Item As Integer


Private Sub Form_Load()
Dim i As Integer
    'Rende i vari bitmax=4095 solo se attualmente=0
    To4095
    '*********************
    '* PERSONALIZZAZIONI *
    '*********************
    
    'Rimozione schede inutili
    For i = 18 To 7 Step -1
        tbsOptions.Tabs.Remove i
    Next
   DoEvents
        If Canale(0).Attivo = True Then
            AggiornaTbs (1)
        End If
   
    
    'DoEvents
    
    If lDebug = True Or Versione = "Poseidon" Then
    
        If Canale(0).Attivo = True Then
            AggiornaTbs (1)
        End If

        Item = tbsOptions.SelectedItem.Index - 1
        AggiornaTbs (tbsOptions.SelectedItem.Index)
    Else
        'Rende invisibili le rimanenti schede
        Frame1.Caption = ""
        tbsOptions.Visible = False
        oAttivo.Visible = False
        tNome.Visible = False
        Label1.Visible = False
        tUnita.Visible = False
        Label2.Visible = False
        Frame2.Visible = False
        Frame3.Visible = False
        cmdApply.Visible = False
        'Sposta i pulsanti salva e carica setup
       With bSalva
            .Left = 3440
            .Top = 500
        End With
        With bLeggi
            .Left = 1440
            .Top = 500
        End With
        With frmOptions2
            .Width = 6264
            .Height = 2676
        End With

        With Frame4
            .Left = 350
            .Top = 1700
        End With
        With Frame5
            .Left = 2280
            .Top = 1700
        End With
        With bIndietro
            .Left = 3240
            .Top = 1850
        End With
        With cmdOK
            .Left = 4080
            .Top = 1850
        End With
        With cmdCancel
            .Left = 5160
            .Top = 1850
        End With
        With Frame1
            .Left = 360
            .Top = 120
            .Height = 1400
        End With
       
    End If
End Sub

Private Sub Form_Paint()
    Item = tbsOptions.SelectedItem.Index - 1
    AggiornaTbs (tbsOptions.SelectedItem.Index)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        'CloseCom
        Me.Hide
        Unload Me
        fMain.Show
    End If
End Sub

Private Sub bIndietro_Click()
    Me.Hide
    fMain.Show
End Sub

Private Sub bSalva_Click()
    Dim Filnb As Integer
    Dim i As Integer
    
    i = tbsOptions.SelectedItem.Index - 1
    Applica (i)

    
    On Error GoTo Annulla
    fMain.CmDialog1.CancelError = True
    'Controlla se si vuole sostituire il file,
    'che la directory eventualmente immessa esista,
    'non prende in considerazione files e directory a sola lettura
    'non mostra la casella sola lettura
    fMain.CmDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + cdlOFNHideReadOnly
    'Filtri di dialogo
    fMain.CmDialog1.Filter = "File Programmazione (*.prg)|*.prg|Tutti i file (*.*)|*.*"
    NewPath (sGetAppPath())
    fMain.CmDialog1.FileName = ""
    If InitDirPrg <> "" Then
        fMain.CmDialog1.InitDir = InitDirPrg
    End If
    fMain.CmDialog1.ShowSave
    On Error GoTo 0
    FileOut = fMain.CmDialog1.FileName
    DoEvents
        
    Me.MousePointer = vbHourglass
    'Salva i dati
    Filnb = FreeFile
    Open FileOut For Output As #Filnb
    Print #Filnb, TestataPrg
    Print #Filnb, Stazione
    For i = 0 To MaxCanali
        Print #Filnb, Canale(i).Nome
        If Canale(i).Attivo = True Then
            Print #Filnb, "True"
        Else
            Print #Filnb, "False"
        End If
        Print #Filnb, Canale(i).UnitaMisura
        Print #Filnb, Canale(i).Bitmin
        Print #Filnb, Canale(i).Bitmax
        Print #Filnb, Str(Canale(i).valMin)
        Print #Filnb, Str(Canale(i).valMax)
        Print #Filnb, Str(Canale(i).valOff)
    Next
    
    
    FileOut = ""
    Me.MousePointer = vbDefault
    
    Close #Filnb
    Exit Sub
Annulla:
    Me.MousePointer = vbDefault
    DoEvents
    'CloseCom
End Sub

Private Sub bLeggi_Click()
    Dim Filnb As Integer
    Dim i As Integer
    Dim Stringa As String
    
    On Error GoTo Annulla
    fMain.CmDialog1.CancelError = True
    'Controlla se si vuole sostituire il file,
    'che la directory eventualmente immessa esista,
    'non prende in considerazione files e directory a sola lettura
    'non mostra la casella sola lettura
    fMain.CmDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + cdlOFNHideReadOnly
    'Filtri di dialogo
    fMain.CmDialog1.Filter = "File Programmazione (*.prg)|*.prg|Tutti i file (*.*)|*.*"
    NewPath (sGetAppPath())
    fMain.CmDialog1.FileName = ""
    If InitDirPrg <> "" Then
        fMain.CmDialog1.InitDir = InitDirPrg
    End If
    fMain.CmDialog1.ShowOpen
    On Error GoTo 0
    FileOut = fMain.CmDialog1.FileName
    DoEvents
        
    Me.MousePointer = vbHourglass
    'Salva i dati
    Filnb = FreeFile
    Open FileOut For Input As #Filnb
    Input #Filnb, Stringa
    If Stringa <> TestataPrg Then
        Messaggio = "ERRORE! " + FileOut + " non � un file di configurazione!"
        MsgBox (Messaggio)
        Me.MousePointer = vbNormal
        Exit Sub
    End If
        Input #Filnb, Stazione
        'frmOptions.tStazione.Text = Stazione
    For i = 0 To MaxCanali
        Input #Filnb, Canale(i).Nome
         Input #Filnb, FileOut
        Canale(i).Attivo = CBool(FileOut)
        Input #Filnb, Canale(i).UnitaMisura
        Input #Filnb, Canale(i).Bitmin
        Input #Filnb, Canale(i).Bitmax
        Input #Filnb, Canale(i).valMin
        Input #Filnb, Canale(i).valMax
        Input #Filnb, Canale(i).valOff
    Next
    
    
    FileOut = ""
    Me.MousePointer = vbDefault
    AggiornaTbs (tbsOptions.SelectedItem.Index)
    Close #Filnb
    Exit Sub
Annulla:
    Me.MousePointer = vbDefault
    DoEvents
    'CloseCom
End Sub

Private Sub cmdApply_Click()
    Dim i As Integer
    Dim Bitmin As Long
    Dim Bitmax As Long
    
    i = tbsOptions.SelectedItem.Index - 1
    Applica (i)
    Exit Sub
    
    Bitmax = Int(Val(tBitMax.Text))
    Bitmin = Int(Val(tBitMin.Text))
    If Bitmax > 65535 Then
        Messaggio = "Il valore massimo dell'ADC" + vbCr + "deve essere inferiore a 65536."
        MsgBox (Messaggio)
        tBitMax.Text = "65535"
        Exit Sub
    End If
    If Bitmin < 0 Then
        Messaggio = "Il valore minimo dell'ADC" + vbCr + "non pu� essere inferiore a zero."
        MsgBox (Messaggio)
        tBitMin.Text = "0"
        Exit Sub
    End If
    
    Canale(i).Bitmax = Bitmax
    Canale(i).Bitmin = Bitmin
    Canale(i).valMax = Val2(tValMax.Text)
    Canale(i).valMin = Val2(tValMin.Text)
    Canale(i).valOff = Val2(tValoff.Text)
    Canale(i).Attivo = oAttivo.Value
    Canale(i).Nome = tNome.Text
    Canale(i).UnitaMisura = tUnita.Text
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    fMain.Show
End Sub

Private Sub CmdOK_Click()
    Dim ok As Boolean
    Dim i As Integer
    cmdApply_Click
    ok = False
    For i = 0 To MaxCanali
        If Canale(i).Attivo = True Then
            ok = True
            Exit For
        End If
    Next
    If ok = False Then
        Messaggio = "Nessun canale attivo!"
        i = MsgBox(Messaggio, vbOKOnly, "Errore!")
        Exit Sub
    End If
    Me.Hide
    fPhCond.Show
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'Gestisce la combinazione di tasti CTRL+TAB per lo
    'spostamento sulla scheda successiva.
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.Index
        If i = tbsOptions.Tabs.count Then
            '� stata raggiunta l'ultima scheda e quindi
            'torna alla scheda 1.
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'Incrementa l'indice della scheda
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
End Sub

Private Sub tbsOptions_Click()
    'On Error Resume Next
    Dim i As Integer
    Applica (Item)
    
    'Questo comando aggiorna il database dei canali

    'Qui i controlli sono sempre gli stessi.
    'E' il contenuto che cambia al variare della scheda
    AggiornaTbs (tbsOptions.SelectedItem.Index)
    Item = tbsOptions.SelectedItem.Index - 1
    Me.Refresh
End Sub

Public Sub AggiornaTbs(Elemento As Integer)
    Dim i As Integer
    i = Elemento - 1
    If lDebug Then Frame1.Caption = "Canale " + Str(i + 1)
    tBitMax.Text = Format(Canale(i).Bitmax, "0")
    tBitMin.Text = Format(Canale(i).Bitmin, "0")
    tValMax.Text = Str(Canale(i).valMax)
    tValMin.Text = Str(Canale(i).valMin)
    tValoff.Text = Str(Canale(i).valOff)
    If Canale(i).Attivo = True Then
        oAttivo.Value = 1
    Else
        oAttivo.Value = 0
    End If
    tNome.Text = Trim(Canale(i).Nome)
    tUnita.Text = Trim(Canale(i).UnitaMisura)
End Sub

Public Sub Applica(scheda As Integer)
    Dim i As Integer
    Dim Bitmin As Long
    Dim Bitmax As Long
    
    i = scheda
    Bitmax = Int(Val(tBitMax.Text))
    Bitmin = Int(Val(tBitMin.Text))
    If Bitmax > 65535 Then
        Messaggio = "Il valore massimo dell'ADC" + vbCr + "deve essere inferiore a 65536."
        MsgBox (Messaggio)
        tBitMax.Text = "65535"
        Exit Sub
    End If
    If Bitmin < 0 Then
        Messaggio = "Il valore minimo dell'ADC" + vbCr + "non pu� essere inferiore a zero."
        MsgBox (Messaggio)
        tBitMin.Text = "0"
        Exit Sub
    End If

    Canale(i).Bitmax = Bitmax
    Canale(i).Bitmin = Bitmin
    Canale(i).valMax = Val2(tValMax.Text)
    Canale(i).valMin = Val2(tValMin.Text)
    Canale(i).valOff = Val2(tValoff.Text)
    Canale(i).Attivo = oAttivo.Value
    Canale(i).Nome = tNome.Text
    Canale(i).UnitaMisura = tUnita.Text

End Sub

Public Sub To4095()
   Dim i As Integer
   For i = 0 To MaxCanali
      If Canale(i).Bitmax = 0 Then
         Canale(i).Bitmax = 4095
      End If
   Next
   
End Sub

