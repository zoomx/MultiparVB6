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
   Caption         =   "Multipar"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2400
   BeginProperty Font 
      Name            =   "AFPalm"
      Size            =   8.25
      Charset         =   2
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   StartUpPosition =   2  'CenterScreen
   Begin IngotWidgetCtl.AFWidget AFWidget1 
      Height          =   255
      Left            =   1200
      OleObjectBlob   =   "fMain.frx":0000
      TabIndex        =   11
      Top             =   1080
      Width           =   255
   End
   Begin IngotToneCtl.AFTone AFTone1 
      Height          =   480
      Left            =   1920
      OleObjectBlob   =   "fMain.frx":001C
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin IngotLabelCtl.AFLabel lTitle 
      Height          =   255
      Left            =   720
      OleObjectBlob   =   "fMain.frx":0041
      TabIndex        =   9
      Top             =   225
      Width           =   825
   End
   Begin IngotGraphicCtl.AFGraphic AFGraphic1 
      Height          =   735
      Left            =   1560
      OleObjectBlob   =   "fMain.frx":0092
      TabIndex        =   8
      Top             =   225
      Width           =   735
   End
   Begin IngotButtonCtl.AFButton bEnd 
      Height          =   255
      Left            =   1560
      OleObjectBlob   =   "fMain.frx":00CA
      TabIndex        =   7
      Top             =   1200
      Width           =   495
   End
   Begin IngotLabelCtl.AFLabel lSn 
      Height          =   135
      Left            =   240
      OleObjectBlob   =   "fMain.frx":0113
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin IngotLabelCtl.AFLabel lData 
      Height          =   135
      Left            =   240
      OleObjectBlob   =   "fMain.frx":015B
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
   End
   Begin IngotLabelCtl.AFLabel lBattery 
      Height          =   135
      Left            =   240
      OleObjectBlob   =   "fMain.frx":01A4
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin IngotButtonCtl.AFButton bRestart 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "fMain.frx":01F1
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin IngotButtonCtl.AFButton bDownload 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "fMain.frx":023D
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin IngotButtonCtl.AFButton bConnect 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "fMain.frx":028A
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin IngotSerialCtl.AFSerial AFSerial1 
      Height          =   480
      Left            =   1920
      OleObjectBlob   =   "fMain.frx":02D6
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
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
        'On Palm use cradle port
        AFSerial1.CommPort = 32768
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
        lData.Caption = lData.Caption + " " + left(Stringa, Len(Stringa) - 2) & " Bytes"
    Else
        GoTo Failed
    End If
    
    Stringa = InputComTimeOut(5)
    If Stringa <> "TimeOut" Then
        'Debug.Print "volts="; Stringa
        'lVolts.Caption = Left(Stringa, Len(Stringa) - 2)
        lBattery.Caption = lBattery.Caption + " " + left(Stringa, Len(Stringa) - 2) & " volt"
        TensioneBatteria = Val2(Stringa)
    Else
        GoTo Failed
    End If

    Stringa = InputComTimeOut(5)
    If Stringa <> "TimeOut" Then
        'Debug.Print "s/n="; Stringa
        'frameCentralina.Caption = "Centralina N. " + Left(Stringa, Len(Stringa) - 2)
        lSn.Caption = lSn.Caption + " " + left(Stringa, Len(Stringa) - 2)
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

