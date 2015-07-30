VERSION 5.00
Begin VB.UserControl MyProgressBar 
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3615
   ScaleHeight     =   495
   ScaleWidth      =   3615
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   600
      Top             =   0
      Width           =   15
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   600
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   550
   End
End
Attribute VB_Name = "MyProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Valori predefiniti proprietà
Const m_def_Value = 0
Const m_def_MinValue = 0
Const m_def_MaxValue = 100

'Variabili Proprietà
Dim m_Value As Long
Dim m_MinValue As Long
Dim m_MaxValue As Long

'Inizializza le proprietà di UserControl
Private Sub UserControl_InitProperties()
    m_Value = m_def_Value
    m_MaxValue = m_def_MaxValue
    m_MinValue = m_def_MinValue
End Sub

Public Property Get Value() As Long
    Value = m_Value
    Dim Ratio As Single
    Ratio = ((m_Value - m_MinValue) / (m_MaxValue - m_MinValue))
    If Ratio > 1 Then Ratio = 1
    Shape2.Width = 15 + Ratio * 3000
    Label1.Caption = Int(Ratio * 100) & "%"
End Property

Public Property Let Value(ByVal New_Value As Long)
    m_Value = New_Value
    PropertyChanged "Value"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Shape2.FillColor = PropBag.ReadProperty("FillColor", &H0&)
    m_MaxValue = PropBag.ReadProperty("MaxValue", m_def_MaxValue)
    m_MinValue = PropBag.ReadProperty("MinValue", m_def_MinValue)
    Label1.Caption = PropBag.ReadProperty("Caption", "")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
    Call PropBag.WriteProperty("FillColor", Shape2.FillColor, &H0&)
    Call PropBag.WriteProperty("MaxValue", m_MaxValue, m_def_MaxValue)
    Call PropBag.WriteProperty("MinValue", m_MinValue, m_def_MinValue)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "")
End Sub

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Restituisce un oggetto Font."
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    PropertyChanged "Font"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=Shape2,Shape2,-1,FillColor
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Restituisce o imposta il colore utilizzato per applicare riempimenti a forme, cerchi e caselle."
    FillColor = Shape2.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    Shape2.FillColor() = New_FillColor
    PropertyChanged "FillColor"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=8,0,0,100
Public Property Get MaxValue() As Long
    MaxValue = m_MaxValue
End Property

Public Property Let MaxValue(ByVal New_MaxValue As Long)
    m_MaxValue = New_MaxValue
    PropertyChanged "MaxValue"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=8,0,0,0
Public Property Get MinValue() As Long
    MinValue = m_MinValue
End Property

Public Property Let MinValue(ByVal New_MinValue As Long)
    m_MinValue = New_MinValue
    PropertyChanged "MinValue"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Restituisce o imposta il testo visualizzato sulla barra del titolo o sotto l'icona di un oggetto."
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

