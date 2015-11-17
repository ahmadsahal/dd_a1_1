VERSION 5.00
Begin VB.Form frmKamus 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Terjemahan Indonesia - Sunda"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Bookman Old Style"
      Size            =   12
      Charset         =   0
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKamus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
      Caption         =   "&Tutup"
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Caption         =   "&Bahasa"
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   6015
      Begin VB.OptionButton optBahasa 
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
      Begin VB.OptionButton optBahasa 
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblSunda 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sunda - Indonesia"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   600
         TabIndex        =   8
         Top             =   360
         Width           =   2220
      End
      Begin VB.Label lblINA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Indonesia - Sunda"
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   3600
         TabIndex        =   7
         Top             =   360
         Width           =   2220
      End
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Clear"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox t3 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox t2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   3120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox t1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   2655
   End
End
Attribute VB_Name = "frmKamus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function _
WritePrivateProfileString Lib _
"kernel32.dll" Alias _
"WritePrivateProfileStringA" ( _
ByVal lpApplicationName As String, _
ByVal lpKeyName As String, _
ByVal lpString As String, _
ByVal lpFileName As String) As Long

Private Declare Function _
GetPrivateProfileString Lib _
"kernel32.dll" Alias _
"GetPrivateProfileStringA" ( _
ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, _
ByVal lpDefault As String, _
ByVal lpReturnedString As String, _
ByVal nSize As Long, _
ByVal lpFileName As String) As Long

Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Activate()
    Dim B1 As String * 255
    Dim Y1: Dim R1, Auto
    MsgBox "Masukkan kata yang ingin " & _
    "diterjemahkan " & vbCrLf & _
    "pada textbox, " & _
    "kemudian tekan spasi", vbInformation, _
    "Kamus Viansastra"
    t1.Text = "": t2.Text = "": t3.Text = ""
    t1.SetFocus
End Sub

Sub UbahKata()
    Dim Kata, i, A
    Dim B1 As String * 255
    Dim Y1, R1, Auto
    Auto = App.Path & "\kic1.kms"
    '
    Kata = t1.Text
    i = InStr(A, Kata)
    A = Mid(A, i + 1)
    '
'**************************************************
If optBahasa(0).Value = True Then
    Y1 = GetPrivateProfileString("KamusIS", _
    CStr(Kata), CStr(Kata), B1, 255, Auto)
    R1 = Left(B1, Y1): Me.t2.SelText = R1 & " "
    t1.SelStart = Len(Kata)
'**************************************************
ElseIf optBahasa(1).Value = True Then
    Y1 = GetPrivateProfileString("KamusSI", _
    CStr(Kata), CStr(Kata), B1, 255, Auto)
    R1 = Left(B1, Y1): Me.t2.SelText = R1 & " "
    t1.SelStart = Len(Kata)
'**************************************************
Else: MsgBox "Pilih alih bahasa", 48
End If
End Sub

Private Sub cmdClear_Click()
    Me.t1.Text = ""
    Me.t2.Text = ""
    Me.t3.Text = ""
    Me.t1.SetFocus
End Sub

Private Sub Form_Load()
    Me.optBahasa(0).Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next: Dim BIKIN, Alamat
BIKIN = WritePrivateProfileString( _
"Akhir", "kalimat", CStr(t2.Text), Alamat)
End Sub

Private Sub lblINA_Click()
    Me.optBahasa(1).Value = True
End Sub

Private Sub lblSunda_Click()
    Me.optBahasa(0).Value = True
End Sub

Private Sub t1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
        UbahKata
        t3.SelText = t1.Text
        t1.Text = ""
    Ahir = t3.SelLength
    End If
End Sub
