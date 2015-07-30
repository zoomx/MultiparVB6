VERSION 5.00
Object = "{BAE83016-CA11-4D8A-BBFA-0AE9863B82DE}#3.0#0"; "IngotLabelCtl.dll"
Object = "{9F4EED48-8EC7-4316-A47D-F6161874E478}#3.0#0"; "IngotButtonCtl.dll"
Object = "{F9885939-2FBB-491F-8EC3-DBC61CCFA7DB}#3.0#0"; "IngotGraphicCtl.dll"
Object = "{84BE8A4A-3F9A-44E9-9B5E-E76D4888BA67}#3.0#0"; "IngotToneCtl.dll"
Object = "{4BA93651-C72B-4D5A-8529-AD4762F41507}#3.0#0"; "IngotSerialCtl.dll"
Object = "{9CC0BC1C-C10B-4F43-90AA-5E3AD0257E45}#3.0#0"; "IngotWidgetCtl.dll"
Begin VB.Form fMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pocket Multipar"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   269
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   240
   StartUpPosition =   2  'CenterScreen
   Begin IngotLabelCtl.AFLabel AFLabel1 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "fMain.frx":0000
      TabIndex        =   12
      Top             =   840
      Width           =   2655
   End
   Begin IngotSerialCtl.AFSerial AFSerial1 
      Height          =   480
      Left            =   2880
      OleObjectBlob   =   "fMain.frx":005E
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin IngotButtonCtl.AFButton bConnect 
      Height          =   375
      Left            =   360
      OleObjectBlob   =   "fMain.frx":009C
      TabIndex        =   10
      Top             =   1320
      Width           =   1095
   End
   Begin IngotButtonCtl.AFButton bDownload 
      Height          =   375
      Left            =   360
      OleObjectBlob   =   "fMain.frx":00E8
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin IngotButtonCtl.AFButton bRestart 
      Height          =   360
      Left            =   1560
      OleObjectBlob   =   "fMain.frx":0135
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
   End
   Begin IngotLabelCtl.AFLabel lBattery 
      Height          =   240
      Left            =   360
      OleObjectBlob   =   "fMain.frx":0181
      TabIndex        =   7
      Top             =   2280
      Width           =   1575
   End
   Begin IngotLabelCtl.AFLabel lData 
      Height          =   240
      Left            =   360
      OleObjectBlob   =   "fMain.frx":01CE
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
   End
   Begin IngotLabelCtl.AFLabel lSn 
      Height          =   240
      Left            =   360
      OleObjectBlob   =   "fMain.frx":0217
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
   End
   Begin IngotButtonCtl.AFButton bEnd 
      Height          =   480
      Left            =   2400
      OleObjectBlob   =   "fMain.frx":025F
      TabIndex        =   4
      Top             =   2760
      Width           =   735
   End
   Begin IngotGraphicCtl.AFGraphic AFGraphic1 
      Height          =   735
      Left            =   2520
      OleObjectBlob   =   "fMain.frx":02A8
      TabIndex        =   3
      Top             =   0
      Width           =   735
   End
   Begin IngotLabelCtl.AFLabel lTitle 
      Height          =   375
      Left            =   600
      OleObjectBlob   =   "fMain.frx":02E0
      TabIndex        =   2
      Top             =   120
      Width           =   1545
   End
   Begin IngotToneCtl.AFTone AFTone1 
      Height          =   480
      Left            =   2880
      OleObjectBlob   =   "fMain.frx":032D
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   480
   End
   Begin IngotWidgetCtl.AFWidget AFWidget1 
      Height          =   255
      Left            =   2400
      OleObjectBlob   =   "fMain.frx":0352
      TabIndex        =   0
      Top             =   1680
      Width           =   255
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sleeper As New Sleeper

Private Sub bConnect_Click()
    Dim Stringa As String

    #If APPFORGE Then
        'On Pocket use cradle port
        AFSerial1.CommPort = 1
    #Else
        'In Windows use the COM1 port
        
        AFSerial1.CommPort = 1
    #End If
    AFSerial1.Settings = "19200,n,8,1"
    OpenCom
    AFSerial1.InBufferCount = 0
    AFSerial1.Output = Chr$(3)
    AFTone1.Duration = 250
    AFTone1.Play
    AFSerial1.Output = Chr$(3)
    AFTone1.Duration = 250
    AFTone1.Play
    AFSerial1.Output = Chr$(3)
    
    Stringa = InputComTimeOut(5)
    If InStr(Stringa, "Poseidon") <> 1 Then
        MsgBox ("Connessione fallita->" & Stringa)
        Exit Sub
    Else
        'MsgBox ("Connessione OK!")
    End If
    
    AFTone1.Duration = 250
    AFTone1.Play
    
    fMain.AFSerial1.InBufferCount = 0
    fMain.AFSerial1.Output = InfoAcq + vbCr
    Stringa = ""
    Stringa = InputComTimeOut(5)
    If Stringa <> "TimeOut" Then
        'Debug.Print "bytes="; Stringa
        'lBytes.Caption = Left(Stringa, Len(Stringa) - 2)
        lData.Caption = lData.Caption + " " + Left(Stringa, Len(Stringa) - 2) & " Bytes"
    Else
        GoTo Failed
    End If
    
    Stringa = InputComTimeOut(5)
    If Stringa <> "TimeOut" Then
        'Debug.Print "volts="; Stringa
        'lVolts.Caption = Left(Stringa, Len(Stringa) - 2)
        lBattery.Caption = lBattery.Caption + " " + Left(Stringa, Len(Stringa) - 2) & " volt"
        TensioneBatteria = Val2(Stringa)
    Else
        GoTo Failed
    End If

    Stringa = InputComTimeOut(5)
    If Stringa <> "TimeOut" Then
        'Debug.Print "s/n="; Stringa
        'frameCentralina.Caption = "Centralina N. " + Left(Stringa, Len(Stringa) - 2)
        lSn.Caption = lSn.Caption + " " + Left(Stringa, Len(Stringa) - 2)
    Else
        GoTo Failed
    End If

    'Legge il Fattore Batteria
    Sleeper.Sleep 250
    fMain.AFSerial1.InBufferCount = 0
    fMain.AFSerial1.Output = LeggiBattFact + vbCr
    Stringa = ""
    Stringa = InputComTimeOut(5)
    If Stringa <> "TimeOut" Then
        FattoreBatteriaInterna = Val2(Stringa)
        'Debug.Print "factBatt="; Stringa
    Else
        GoTo Failed
    End If

    'Risposta = ScaricaProgrammazione
    'If Risposta <> 0 Then ProgrammazioneCaricata = True
 
    Me.AFSerial1.InBufferCount = 0
    'Me.MousePointer = vbNormal
    AbilitaTasti
    Me.AFSerial1.InBufferCount = 0
    AFTone1.Duration = 350
    AFTone1.Play

    
    Exit Sub

Failed:
        Stringa = "Errore nella comunicazione" + vbCr + "il datalogger non risponde!" + vbCr + "Controllare che sia in modo Comandi" + vbCr + "  Controllare il cavo di collegamento"
        MsgBox (Stringa)

        CloseCom
    

End Sub
Private Sub bDownload_Click()
    FileOut = ""
    FileOut = Trim(Str(Year(Now)))
    Stringa = Trim(Str(Month(Now)))
    If Len(Stringa) = 1 Then Stringa = "0" + Stringa
    FileOut = FileOut + Stringa
    Stringa = Trim(Str(Day(Now)))
    If Len(Stringa) = 1 Then Stringa = "0" + Stringa
    FileOut = FileOut + Stringa
    Stringa = Trim(Str(Hour(Now)))
    If Len(Stringa) = 1 Then Stringa = "0" + Stringa
    FileOut = FileOut + Stringa
    Stringa = Trim(Str(Minute(Now)))
    If Len(Stringa) = 1 Then Stringa = "0" + Stringa
    FileOut = FileOut + Stringa
    Stringa = Trim(Str(Second(Now)))
    If Len(Stringa) = 1 Then Stringa = "0" + Stringa

    FileOut = FileOut + ".dat"
    FileOut = InputBox("File name?", "Save", FileOut)
    If FileOut = "" Then Exit Sub
    #If APPFORGE Then
        FileOut = "Card:\" + FileOut + Stringa + ".txt" '+ Location + ".txt"
    #Else
        FileOut = FileOut + Stringa + ".txt" '+ Location + ".txt"
    #End If

End Sub

Private Sub bEnd_Click()
    CloseCom
    Unload Me
    End
End Sub
Public Sub AbilitaTasti()
    'Abilita i tasti del form principale
    bDownload.Enabled = True
    bRestart.Enabled = True
    DoEvents
End Sub

Public Sub DisabTasti()
    'Disabilita i tasti del form principale
    bDownload.Enabled = False
    bRestart.Enabled = False
    DoEvents
End Sub


