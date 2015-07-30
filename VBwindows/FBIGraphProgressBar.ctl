VERSION 5.00
Begin VB.UserControl FBIGraphProgressBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3210
   FillColor       =   &H00FF0000&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   MaskColor       =   &H80000008&
   ScaleHeight     =   645
   ScaleWidth      =   3210
   ToolboxBitmap   =   "FBIGraphProgressBar.ctx":0000
End
Attribute VB_Name = "FBIGraphProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' FBI Graph Progress Bar
' Copyright Fibia FBI - Team 18/VBSimple - Marzo 2002
'
' Struttura del codice:
'  |-> Dichiarazioni    (Private)
'  |-> Eventi           (Private)
'  |-> Funzioni Interne (Private)
'  |-> Proprietà        (Public)
'  \-> Funzioni Esterne (Public)

'------------------- DICHIARAZIONI ---------------------

Private lngMinimo As Long
Private lngMassimo As Long
Private lngValore As Long
Private lngBackColor As Long
Private lngForeColor As Long
Private lngFillColor As Long
Private Allineamento As AlignmentConstants
Private blnDefaultColorMode As Boolean
Private strFormat As String

Public Event Changed(ByVal oldValue As Long)
Attribute Changed.VB_Description = "Generato ogni volta che il valore della barra viene modificato."
Public Event Click()
Attribute Click.VB_Description = "Viene generato quando si preme e quindi si rilascia un pulsante del mouse su un oggetto."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Viene generato quando si preme e si rilascia due volte in rapida successione un pulsante del mouse su un oggetto."
Attribute DblClick.VB_UserMemId = -601
Public Event Error(ByVal Errore As String, ByRef Cancel As Boolean)
Attribute Error.VB_Description = "Viene generato quando si verifica un errore."
Public Event Massimo()
Attribute Massimo.VB_Description = "Viene generato quando il valore della barra raggiunge il valore massimo."
Public Event Minimo()
Attribute Minimo.VB_Description = "Viene generato quando il valore della barra raggiunge il valore minimo."
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Attribute MouseDown.VB_Description = "Viene generato quando si preme il pulsante del mouse mentre lo stato attivo si trova su un oggetto."
Attribute MouseDown.VB_UserMemId = -605
Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Attribute MouseMove.VB_Description = "Viene generato quando si sposta il mouse."
Attribute MouseMove.VB_UserMemId = -606
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Attribute MouseUp.VB_Description = "Viene generato quando si rilascia il pulsante del mouse mentre lo stato attivo si trova su un oggetto."
Attribute MouseUp.VB_UserMemId = -607

'---------------------- EVENTI -------------------------

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_InitProperties()
    blnDefaultColorMode = True
    Allineamento = vbCenter
    lngBackColor = &H8000000F
    lngFillColor = &H8000000D
    lngForeColor = &H8000000E
    lngMinimo = 0
    lngMassimo = 100
    lngValore = 0
    strFormat = "$V"
    CambiaColorMode
    UserControl.AutoRedraw = True
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Allineamento = .ReadProperty("Alignment", vbCenter)
        lngBackColor = .ReadProperty("BackColor", &HFFFFFF)
        blnDefaultColorMode = .ReadProperty("DefaultColorMode", True)
        lngFillColor = .ReadProperty("FillColor", &HFF0000)
        lngForeColor = .ReadProperty("ForeColor", &HFF0000)
        strFormat = .ReadProperty("Format", "$V")
        Set UserControl.Font = .ReadProperty("Font")
        lngMassimo = .ReadProperty("Max", 100)
        lngMinimo = .ReadProperty("Min", 0)
        Set UserControl.MouseIcon = .ReadProperty("MouseIcon")
        UserControl.MousePointer = .ReadProperty("MousePointer", vbDefault)
        lngValore = .ReadProperty("Value", 0)
    End With
    CambiaColorMode
    AggiornaGraficamente
End Sub

Private Sub UserControl_Resize()
    If UserControl.Width < 135 Then UserControl.Width = 135
    If UserControl.Height < 75 Then UserControl.Height = 75
    AggiornaGraficamente
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Alignment", Allineamento, vbCenter)
        Call .WriteProperty("BackColor", lngBackColor, &HFFFFFF)
        Call .WriteProperty("DefaultColorMode", blnDefaultColorMode, True)
        Call .WriteProperty("FillColor", lngFillColor, &HFF0000)
        Call .WriteProperty("ForeColor", lngForeColor, &HFF0000)
        Call .WriteProperty("Format", strFormat, "$V")
        Call .WriteProperty("Font", UserControl.Font)
        Call .WriteProperty("Max", lngMassimo, 100)
        Call .WriteProperty("Min", lngMinimo, 0)
        Call .WriteProperty("MouseIcon", UserControl.MouseIcon)
        Call .WriteProperty("MousePointer", UserControl.MousePointer, vbDefault)
        Call .WriteProperty("Value", lngValore, 0)
    End With
End Sub

'----------------- FUNZIONI INTERNE --------------------

Private Sub AggiornaGraficamente()
    Dim Testo As String
    Dim PosX As Single
    Dim PosY As Single
    With UserControl
        Testo = Caption
        .Cls
        Select Case Allineamento
            Case vbLeftJustify = 15
            Case vbCenter: PosX = (.ScaleWidth - 15 - .TextWidth(Testo)) / 2
            Case vbRightJustify: PosX = .ScaleWidth - 15 - .TextWidth(Testo)
        End Select
        
        PosY = (.ScaleHeight - 15 - .TextHeight(Testo)) / 2
        If blnDefaultColorMode = True Then
            .DrawMode = vbCopyPen
            .CurrentX = PosX
            .CurrentY = PosY
            UserControl.Print Testo
            .DrawMode = vbMergePenNot
            If lngValore > 0 Then UserControl.Line (0, 0)-(.ScaleWidth * (lngValore - lngMinimo) / (lngMassimo - lngMinimo), .Height), .FillColor, BF
        Else
            .DrawMode = vbCopyPen
            If lngValore > 0 Then UserControl.Line (0, 0)-(.ScaleWidth * (lngValore - lngMinimo) / (lngMassimo - lngMinimo), .Height), .FillColor, BF
            .CurrentX = PosX
            .CurrentY = PosY
            UserControl.Print Testo
        End If
    End With
End Sub

Private Sub CambiaColorMode()
    With UserControl
        If blnDefaultColorMode = True Then
            .BackColor = &HFFFFFF
            .FillColor = &HFF0000
            .ForeColor = &HFF0000
        Else
            .FillColor = lngFillColor
            .BackColor = lngBackColor
            .ForeColor = lngForeColor
        End If
    End With
End Sub

Private Sub GestisciErrore(ByVal Messaggio As String)
    Dim Cancel As Boolean
    Cancel = False
    RaiseEvent Error(Messaggio, Cancel)
    If Cancel = False Then MsgBox Messaggio, vbCritical Or vbOKOnly, "FBIGraphProgressBar"
End Sub

Private Function Replace(sIn As String, sFind As String, sReplace As String, Optional nStart As Long = 1, Optional nCount As Long = -1, Optional bCompare As VbCompareMethod = vbBinaryCompare) As String
    Dim nC As Long, nPos As Integer, sOut As String
    sOut = sIn
    nPos = InStr(nStart, sOut, sFind, bCompare)
    If nPos = 0 Then GoTo EndFn:
    Do
        nC = nC + 1
        sOut = Left(sOut, nPos - 1) & sReplace & Mid(sOut, nPos + Len(sFind))
        If nCount <> -1 And nC >= nCount Then Exit Do
        nPos = InStr(nStart, sOut, sFind, bCompare)
    Loop While nPos > 0
EndFn:
    Replace = sOut
End Function

'--------------------- PROPRIETA' ----------------------

Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "Restituisce o imposta il valore dell'allineamento del testo all'interno della barra di avanzamento."
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Aspetto"
    Alignment = Allineamento
End Property

Public Property Let Alignment(ByVal newAlignment As AlignmentConstants)
    Allineamento = newAlignment
    PropertyChanged "Alignment"
    AggiornaGraficamente
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Restituisce o imposta il colore di sfondo utilizzato per la visualizzazione di testo e grafica in un oggetto."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Aspetto"
Attribute BackColor.VB_UserMemId = -501
    BackColor = lngBackColor
End Property

Public Property Let BackColor(ByVal newBackColor As OLE_COLOR)
    lngBackColor = newBackColor
    If blnDefaultColorMode = False Then UserControl.BackColor = newBackColor
    PropertyChanged "BackColor"
    AggiornaGraficamente
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Restituisce il testo contenuto nella barra di avanzamento."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Testo"
    Caption = strFormat
    Caption = Replace(Caption, "$m", CStr(lngMinimo), , , vbBinaryCompare)
    Caption = Replace(Caption, "$M", CStr(lngMassimo), , , vbBinaryCompare)
    Caption = Replace(Caption, "$V", CStr(lngValore), , , vbTextCompare)
    Caption = Replace(Caption, "$P", CStr(Int((lngValore - lngMinimo) / (lngMassimo - lngMinimo) * 100)) & "%", , , vbTextCompare)
End Property

Public Property Get DefaultColorMode() As Boolean
Attribute DefaultColorMode.VB_Description = "Restituisce o imposta un valore corrispondente alla modalità di colorazione della barra."
Attribute DefaultColorMode.VB_ProcData.VB_Invoke_Property = ";Aspetto"
    DefaultColorMode = blnDefaultColorMode
End Property

Public Property Let DefaultColorMode(ByVal newDefaultColorMode As Boolean)
    blnDefaultColorMode = newDefaultColorMode
    CambiaColorMode
    PropertyChanged "DefaultColorMode"
    AggiornaGraficamente
End Property

Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Restituisce o imposta il colore utilizzato per applicare riempimenti a forme, cerchi e caselle."
Attribute FillColor.VB_ProcData.VB_Invoke_Property = ";Aspetto"
Attribute FillColor.VB_UserMemId = -510
    FillColor = lngFillColor
End Property

Public Property Let FillColor(ByVal newFillColor As OLE_COLOR)
    lngFillColor = newFillColor
    If blnDefaultColorMode = False Then UserControl.FillColor = newFillColor
    PropertyChanged "FillColor"
    AggiornaGraficamente
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Restituisce ed imposta il tipo di carattere utilizzato per la Caption della barra di avanzamento."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Carattere"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal newFont As StdFont)
    Set UserControl.Font = newFont
    PropertyChanged "Font"
    AggiornaGraficamente
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Restituisce o imposta il colore di primo piano utilizzato per la visualizzazione di testo e grafica in un oggetto."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Aspetto"
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = lngForeColor
End Property

Public Property Let ForeColor(ByVal newForeColor As OLE_COLOR)
    lngForeColor = newForeColor
    If blnDefaultColorMode = False Then UserControl.ForeColor = lngForeColor
    PropertyChanged "ForeColor"
    AggiornaGraficamente
End Property

Public Property Get Format() As String
Attribute Format.VB_Description = "Definisce la regola per la visualizzazione del testo nella barra di avanzamento."
Attribute Format.VB_ProcData.VB_Invoke_Property = ";Testo"
    Format = strFormat
End Property

Public Property Let Format(ByVal newFormat As String)
    strFormat = newFormat
    PropertyChanged "Format"
    AggiornaGraficamente
End Property

Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Restituisce un handle fornito in Microsoft Windows al contesto di periferica di un oggetto."
    hDC = UserControl.hDC
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Restituisce un handle (da Microsoft Windows) alla finestra di un oggetto."
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

Public Property Get Max() As Long
Attribute Max.VB_Description = "Restituisce ed imposta il valore massimo per la barra di avanzamento."
Attribute Max.VB_ProcData.VB_Invoke_Property = ";Dati"
    Max = lngMassimo
End Property

Public Property Let Max(ByVal newMax As Long)
    Dim oldValue As Long
    If newMax < lngMinimo Then
        GestisciErrore "Errore nell'impostazione del valore"
        Exit Property
    End If
    lngMassimo = newMax
    If newMax < lngValore Then
        oldValue = lngValore
        lngValore = newMax
        RaiseEvent Changed(oldValue)
        RaiseEvent Massimo
    End If
    PropertyChanged "Max"
    AggiornaGraficamente
End Property

Public Property Get Min() As Long
Attribute Min.VB_Description = "Restituisce ed imposta il valore minimo per la barra di avanzamento."
Attribute Min.VB_ProcData.VB_Invoke_Property = ";Dati"
    Min = lngMinimo
End Property

Public Property Let Min(ByVal newMin As Long)
    Dim oldValue As Long
    If newMin > lngMassimo Then
        GestisciErrore "Errore nell'impostazione del valore"
        Exit Property
    End If
    lngMinimo = newMin
    If newMin > lngValore Then
        oldValue = lngValore
        lngValore = newMin
        RaiseEvent Changed(oldValue)
        RaiseEvent Minimo
    End If
    PropertyChanged "Min"
    AggiornaGraficamente
End Property

Public Property Get MouseIcon() As StdPicture
Attribute MouseIcon.VB_Description = "Imposta un'icona personalizzata per il puntatore del mouse."
Attribute MouseIcon.VB_ProcData.VB_Invoke_Property = "StandardPicture;Aspetto"
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal newMouseIcon As StdPicture)
    Set UserControl.MouseIcon = newMouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Restituisce o imposta il tipo di puntatore del mouse visualizzato quando il puntatore si trova su una parte specifica di un oggetto."
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";Aspetto"
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal newMousePointer As MousePointerConstants)
    UserControl.MousePointer() = newMousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get Value() As Long
Attribute Value.VB_Description = "Restituisce o imposta il valore corrente della proprietà Value di un controllo."
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Dati"
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    Value = lngValore
End Property

Public Property Let Value(ByVal newValue As Long)
    Dim oldValue As Long
    If (newValue < lngMinimo) Or (newValue > lngMassimo) Then
        GestisciErrore "Errore nell'impostazione del valore"
        Exit Property
    End If
    oldValue = lngValore
    lngValore = newValue
    RaiseEvent Changed(oldValue)
    If lngValore = lngMinimo Then RaiseEvent Minimo
    If lngValore = lngMassimo Then RaiseEvent Massimo
    PropertyChanged "Value"
    AggiornaGraficamente
End Property

'----------------- FUNZIONI ESTERNE --------------------

Public Sub AboutBox()
Attribute AboutBox.VB_Description = "Informazioni sul controllo."
Attribute AboutBox.VB_UserMemId = -552
    MsgBox "FBI Graph Progress Bar - Copyright Fibia FBI Marzo 2002" & vbNewLine & _
        "Una semplicissima barra di avanzamento con l'effetto Xor " & _
        "grafico tipico dei programmi di installazione, nel rispetto " & _
        "del piano ZeroOCX." & vbNewLine, vbInformation Or vbOKOnly, _
        "FBIGraphProgressBar"
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Aggiorna l'aspetto grafico del controllo."
Attribute Refresh.VB_UserMemId = -550
    AggiornaGraficamente
End Sub
