VERSION 5.00
Begin VB.UserControl TextCurrency 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "TextCurrency.ctx":0000
   Begin VB.TextBox txtOcx 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Text            =   "123,456,789"
      Top             =   600
      Width           =   3975
   End
End
Attribute VB_Name = "TextCurrency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Event Ubah()
Event Klik()
Event KlikGanda()
Event Terfokus()
Event HilangFokus()
Event TekanTombol(VianAscii As Integer)
Event TombolBawah(VianCode As Integer, _
Shift As Integer)
Event TombolAtas(VianCode As Integer, _
Shift As Integer)
Event MouseBawah(VianButton As Integer, _
Shift As Integer, X As Single, Y As Single)
Event MousePindah(VianButton As Integer, _
Shift As Integer, X As Single, Y As Single)
Event MouseAtas(VianButton As Integer, _
Shift As Integer, X As Single, Y As Single)
Event Validasi(Cancel As Boolean)

Public Property Get WarnaDasar() As OLE_COLOR
    WarnaDasar = txtOcx.BackColor
End Property

Public Property Let WarnaDasar( _
ByVal New_BackColor As OLE_COLOR)
    txtOcx.BackColor() = New_BackColor
    PropertyChanged "WarnaDasar"
End Property

Public Property Get GayaBingkai() As Integer
    GayaBingkai = txtOcx.BorderStyle
End Property

Public Property Let GayaBingkai( _
ByVal New_BorderStyle As Integer)
    txtOcx.BorderStyle() = New_BorderStyle
    PropertyChanged "GayaBingkai"
End Property

Private Sub txtOcx_Click()
    RaiseEvent Klik
End Sub

Public Property Get DataMember() As String
    DataMember = txtOcx.DataMember
End Property

Public Property Let DataMember( _
ByVal New_DataMember As String)
    txtOcx.DataMember() = New_DataMember
    PropertyChanged "DataMember"
End Property

Private Sub txtOcx_DblClick()
    RaiseEvent KlikGanda
End Sub

Public Property Get Keaktifan() As Boolean
    Keaktifan = txtOcx.Enabled
End Property

Public Property Let Keaktifan(ByVal _
New_Enabled As Boolean)
    txtOcx.Enabled() = New_Enabled
    PropertyChanged "Keaktifan"
End Property

Public Property Get Huruf() As Font
    Set Huruf = txtOcx.Font
End Property

Public Property Set Huruf(ByVal _
New_Font As Font)
    Set txtOcx.Font = New_Font
    PropertyChanged "Huruf"
End Property

Public Property Get WarnaHuruf() As OLE_COLOR
    WarnaHuruf = txtOcx.ForeColor
End Property

Public Property Let WarnaHuruf( _
ByVal New_ForeColor As OLE_COLOR)
    txtOcx.ForeColor() = New_ForeColor
    PropertyChanged "WarnaHuruf"
End Property

Private Sub txtOcx_KeyDown(KeyCode As Integer, _
Shift As Integer)
    RaiseEvent TombolBawah(KeyCode, Shift)
End Sub

Private Sub txtOcx_KeyUp(KeyCode As Integer, _
Shift As Integer)
    RaiseEvent TombolAtas(KeyCode, Shift)
End Sub

Public Property Get Kunci() As Boolean
    Kunci = txtOcx.Locked
End Property

Public Property Let Kunci(ByVal New_Locked _
As Boolean)
    txtOcx.Locked() = New_Locked
    PropertyChanged "Kunci"
End Property

Public Property Get PanjangMax() As Long
    PanjangMax = txtOcx.MaxLength
End Property

Public Property Let PanjangMax(ByVal _
New_MaxLength As Long)
    txtOcx.MaxLength() = New_MaxLength
    PropertyChanged "PanjangMax"
End Property

Private Sub txtOcx_MouseDown(Button As _
Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseBawah(Button, Shift, X, Y)
End Sub

Public Property Get MouseIcon() As Picture
    Set MouseIcon = txtOcx.MouseIcon
End Property

Public Property Set MouseIcon(ByVal _
 New_MouseIcon As Picture)
    Set txtOcx.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub txtOcx_MouseMove(Button _
As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MousePindah(Button, Shift, X, Y)
End Sub

Public Property Get MousePointer() As Integer
    MousePointer = txtOcx.MousePointer
End Property

Public Property Let MousePointer(ByVal _
New_MousePointer As Integer)
    txtOcx.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub txtOcx_MouseUp(Button _
As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseAtas(Button, Shift, X, Y)
End Sub

Public Property Get MultiLine() As Boolean
    MultiLine = txtOcx.MultiLine
End Property

Public Property Get Teks() As String
    Teks = txtOcx.Text
End Property

Public Property Let Teks(ByVal New_Text _
As String)
    txtOcx.Text() = New_Text
    PropertyChanged "Teks"
End Property

Private Sub txtOcx_Validate(Cancel As Boolean)
    RaiseEvent Validasi(Cancel)
End Sub

Public Property Get ToolTipText() As String
    ToolTipText = txtOcx.ToolTipText
End Property

Public Property Let ToolTipText( _
ByVal New_ToolTipText As String)
    txtOcx.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Private Sub UserControl_ReadPropertyes( _
PropBag As PropertyBag)
    txtOcx.BackColor = PropBag.ReadProperty( _
"WarnaDasar", &H80000005)
    txtOcx.BorderStyle = PropBag.ReadProperty _
("GayaBingkai", 1)
    Set DataFormat = PropBag.ReadProperty( _
"DataFormat", Nothing)
    txtOcx.DataMember = PropBag.ReadProperty( _
"DataMember", "")
    Set DataSource = PropBag.ReadProperty( _
"DataSource", Nothing)
    txtOcx.Enabled = PropBag.ReadProperty( _
"Keaktifan", True)
    Set txtOcx.Font = PropBag.ReadProperty( _
"Huruf", Ambient.Font)
    txtOcx.ForeColor = PropBag.ReadProperty( _
"WarnaHuruf", &H80000008)
    txtOcx.Locked = PropBag.ReadProperty( _
"Kunci", False)
    txtOcx.MaxLength = PropBag.ReadProperty( _
"PanjangMax", 0)
    Set MouseIcon = PropBag.ReadProperty( _
"MouseIcon", Nothing)
    txtOcx.MousePointer = PropBag.ReadProperty _
("MousePointer", 0)
 txtOcx.Text = PropBag.ReadProperty( _
"Teks", "123,456,789")
    txtOcx.ToolTipText = PropBag.ReadProperty( _
"ToolTipText", "")
End Sub

'Write Property values to storage
Private Sub UserControl_WritePropertyes( _
PropBag As PropertyBag)
    Call PropBag.WriteProperty( _
"WarnaDasar", txtOcx.BackColor, &H80000005)
    Call PropBag.WriteProperty( _
"GayaBingkai", txtOcx.BorderStyle, 1)
    Call PropBag.WriteProperty( _
"DataFormat", DataFormat, Nothing)
    Call PropBag.WriteProperty( _
"DataMember", txtOcx.DataMember, "")
    Call PropBag.WriteProperty( _
"DataSource", DataSource, Nothing)
    Call PropBag.WriteProperty( _
"Keaktifan", txtOcx.Enabled, True)
    Call PropBag.WriteProperty( _
"Huruf", txtOcx.Font, Ambient.Font)
    Call PropBag.WriteProperty( _
"WarnaHuruf", txtOcx.ForeColor, &H80000008)
    Call PropBag.WriteProperty( _
"Kunci", txtOcx.Locked, False)
    Call PropBag.WriteProperty( _
"PanjangMax", txtOcx.MaxLength, 0)
    Call PropBag.WriteProperty( _
"MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty( _
"MousePointer", txtOcx.MousePointer, 0)
   Call PropBag.WriteProperty( _
"Teks", txtOcx.Text, "123,456,789")
    Call PropBag.WriteProperty( _
"ToolTipText", txtOcx.ToolTipText, "")
End Sub

Private Sub txtOcx_Change()
    RaiseEvent Ubah
    txtOcx.SelStart = Len(txtOcx.Text)
    txtOcx.Text = Format(txtOcx, _
    "###,###,###,###,###,###,###,##0")
End Sub

Private Sub txtOcx_GotFocus()
    RaiseEvent Terfokus
End Sub

Private Sub txtOcx_KeyPress(KeyAscii As Integer)
    RaiseEvent TekanTombol(KeyAscii)
    If Not (KeyAscii >= Asc("0") And _
        KeyAscii <= Asc("9") Or KeyAscii = _
        vbKeyBack Or KeyAscii = 13) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtOcx_LostFocus()
    RaiseEvent HilangFokus
End Sub

Private Sub UserControl_Resize()
    txtOcx.Move 0, 0, Width, Height
End Sub


