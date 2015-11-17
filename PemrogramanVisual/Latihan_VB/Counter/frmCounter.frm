VERSION 5.00
Begin VB.Form frmCounter 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CREATED BY VIANSASTRA"
   ClientHeight    =   1950
   ClientLeft      =   240
   ClientTop       =   1680
   ClientWidth     =   5550
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   DrawWidth       =   5
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   21.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCounter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      DrawStyle       =   1  'Dash
      DrawWidth       =   10
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1560
      ScaleHeight     =   705
      ScaleWidth      =   2745
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   600
      Top             =   1200
   End
   Begin VB.PictureBox pic2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DrawStyle       =   5  'Transparent
      DrawWidth       =   2
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   5520
      TabIndex        =   0
      Top             =   1665
      Width           =   5550
   End
   Begin VB.Label lblStart 
      AutoSize        =   -1  'True
      Caption         =   "lblStart"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblFinish 
      AutoSize        =   -1  'True
      Caption         =   "lblFinish"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label lblJumlah 
      AutoSize        =   -1  'True
      Caption         =   "lblJumlah"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Menu mnuH 
      Caption         =   "harga"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu HFG 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSewa 
         Caption         =   "&Harga Sewa"
      End
      Begin VB.Menu SSS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Tutup"
      End
   End
End
Attribute VB_Name = "frmCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function _
WritePrivateProfileString _
Lib "kernel32" Alias _
"WritePrivateProfileStringA" _
( _
ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, _
ByVal lpString As Any, _
ByVal lpFileName As String _
) As Long

Private Declare Function _
GetPrivateProfileString _
Lib "kernel32" Alias _
"GetPrivateProfileStringA" _
( _
ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, _
ByVal lpDefault As String, _
ByVal lpReturnedString As String, _
ByVal nSize As Long, _
ByVal lpFileName As String _
) As Long

Private Detik%
Private Menit%
Private jam%
Private Harga
Private HRatio
Private Waktu
Private Waktu2
Private JRatio
Private Pembanding
Private HTulis$
Private Tulis$

Dim B2 As String * 255: Dim Y2: Dim R2
Dim Auto$

Sub MulaiHitung()
    Me.lblStart.Caption = Format( _
    Now, "hh:mm:ss")
    Me.Timer1.Enabled = True
End Sub

Private Sub Form_Activate()
    Form_Paint
End Sub

Sub Form_KeyDown(KeyCode As Integer, _
Shift As Integer)
Static K$, L$
    Select Case KeyCode
        Case vbKeyM
            K = K & "M"
        Case vbKeyU
            K = K & "U"
        Case vbKeyL
            K = K & "L"
        Case vbKeyA
            K = K & "A"
        Case vbKeyI
            K = K & "I"
        Case Else: K = ""
    End Select
        If Len(K) > 5 Then K = Right(K, 5)
        If K = "MULAI" Then Me.MulaiHitung
    Select Case KeyCode
        Case vbKeyS
            L = L & "S"
        Case vbKeyT
            L = L & "T"
        Case vbKeyO
            L = L & "O"
        Case vbKeyP
            L = L & "P"
        Case Else: L = ""
    End Select
        If Len(L) > 4 Then L = Right(L, 4)
        If L = "STOP" Then
            Me.Timer1.Enabled = False
        End If
End Sub

Sub Form_Load()
    KeyPreview = True
    '
    Auto = App.Path & "\Counter.ini"
    pic.CurrentX = -10
    pic.CurrentY = -10

    Me.pic.ForeColor = vbRed
    Me.pic.Font.Bold = True
    pic.Print Waktu
    
    pic.CurrentX = 0 - 40
    pic.CurrentY = 0 - 40

    Me.pic.ForeColor = vbYellow
    Me.pic.Font.Bold = True
    pic.Print Waktu
    Dim XStop
    '
    Y2 = GetPrivateProfileString( _
    "Auto", "HargaSewa", "1000", B2, 255, Auto)
    R2 = Left(B2, Y2)
    HRatio = Val(R2)
    Pembanding = HRatio / 3600
    '
    pic.Cls
End Sub

Sub Form_Paint()
Dim T, L, Y
    T = Height
    L = Width
    
    For Y = 0 To T - (pic2.Height + 60)
        Me.FillColor = _
        RGB(255 - _
        (Y * 255 \ T), 500, 500)
        Me.Line _
        (-1, Y - 1)- _
        (L, Y + 1), , B
    Next Y
        
    Me.CurrentX = 60
    Me.CurrentY = -100
    Me.Font.Size = 25
    Me.Font.Bold = True
    Me.ForeColor = QBColor(0)
    Print "Counter"
    
    Me.Font.Size = 25
    Me.CurrentX = 10
    Me.CurrentY = -80
    Me.ForeColor = &HFF8080
    Print "Counter"
    
    Me.CurrentX = 0
    Me.CurrentY = 500
    Me.Font.Size = 10
    Me.Font.Bold = True
    Me.ForeColor = vbBlack
    Print String(100, "#")
    
    Me.CurrentX = 20
    Me.CurrentY = 520
    Me.Font.Size = 10
    Me.ForeColor = vbCyan
    Print String(100, "#")
End Sub

Private Sub Timer1_Timer():
    Dim X, Y, Bikin
    Dim CWaktu
    
    Me.lblFinish.Caption = _
    Format(Time, "hh:mm:ss")
    
    Waktu = Time - TimeValue( _
    Me.lblStart.Caption)
    
    pic.Cls
    
    X = 120
    Y = 0
    Me.pic.CurrentX = X
    Me.pic.CurrentY = Y
    
    Me.pic.ForeColor = QBColor(0)
    Me.pic.Font.Bold = True
    
    pic.Print Format(Waktu, "hh:mm:ss")
    pic2.BorderStyle = 1
    pic.CurrentX = X - 40
    pic.CurrentY = Y - 40
    Me.pic.ForeColor = vbWhite
    Me.pic.Font.Bold = True
    pic.Print Format(Waktu, "hh:mm:ss")
    pic2.Cls
    pic2.CurrentX = 0
    pic2.CurrentY = 0
    
    Me.Caption = Format(Waktu, "hh:mm:ss")
    On Error Resume Next
    Detik = Second(Waktu)
    Menit = Minute(Waktu) * 60
    jam = Hour(Waktu) * 3600
    Waktu2 = (jam + Menit + Detik) _
    * Pembanding
    pic2.CurrentX = 40
    pic2.CurrentY = 0
    pic2.FontBold = True
    pic2.FontSize = 10
    '
    pic2.Print " Terhitung : " & _
    "Rp. " & CStr(Format(Round(Waktu2), _
    "##,##0")) & " (" & Tulis & ")"
    HTulis = CStr(Format(Round(Waktu2), _
    "##,##0"))
    Me.pic.ToolTipText = _
    "Harga sewa Anda = Rp. " & HTulis & _
    " (" & Tulis & ")"
    Terbilang
End Sub

Sub Terbilang()
Dim Angka, A, B, C, D, E, F, G, H, I
Dim A1, A2, A3, A5, A6, A7, A8, A9, A10, A11
Dim Huruf
HTulis = CStr(Format(Round(Waktu2), "##,##0"))
    Angka = HTulis
    A1 = Right(Angka, 1):
    A5 = Right(Angka, 5)
    A2 = Right(Angka, 2):
    A6 = Right(Angka, 6)
    A3 = Right(Angka, 3):
    A7 = Right(Angka, 7)
    A9 = Right(Angka, 9):
    A10 = Right(Angka, 10)
    A11 = Right(Angka, 11):
    Select Case A1
        Case "1": A = "Satu "
        Case "2": A = "Dua "
        Case "3": A = "Tiga "
        Case "4": A = "Empat "
        Case "5": A = "Lima "
        Case "6": A = "Enam "
        Case "7": A = "Tujuh "
        Case "8": A = "Delapan "
        Case "9": A = "Sembilan "
    End Select
    Select Case A2
        Case "10": A = "Sepuluh "
        Case "11": A = "Sebelas "
        Case "12": A = "Dua Belas "
        Case "13": A = "Tiga Belas "
        Case "14": A = "Empat Belas "
        Case "15": A = "Lima Belas "
        Case "16": A = "Enam Belas "
        Case "17": A = "Tujuh Belas "
        Case "18": A = "Delapan Belas "
        Case "19": A = "Sembilan Belas "
    End Select
    Select Case Mid(A2, 1, 1)
        Case "2": B = "Dua Puluh "
        Case "3": B = "Tiga Puluh "
        Case "4": B = "Empat Puluh "
        Case "5": B = "Lima Puluh "
        Case "6": B = "Enam Puluh "
        Case "7": B = "Tujuh Puluh "
        Case "8": B = "Delapan Puluh "
        Case "9": B = "Sembilan Puluh "
        End Select
    Select Case Mid(A3, 1, 1)
        Case "1": C = "Seratus "
        Case "2": C = "Dua Ratus "
        Case "3": C = "Tiga Ratus "
        Case "4": C = "Empat Ratus "
        Case "5": C = "Lima Ratus "
        Case "6": C = "Enam Ratus "
        Case "7": C = "Tujuh Ratus "
        Case "8": C = "Delapan Ratus "
        Case "9": C = "Sembilan Ratus "
    End Select
    Select Case Mid(A5, 1, 1) And _
    Mid(A6, 1, 1) <> "1"
        Case "2": D = "Dua Ribu "
        Case "3": D = "Tiga Ribu "
        Case "4": D = "Empat Ribu "
        Case "5": D = "Lima Ribu "
        Case "6": D = "Enam Ribu "
        Case "7": D = "Tujuh Ribu "
        Case "8": D = "Delapan Ribu "
        Case "9": D = "Sembilan Ribu "
    End Select
    Select Case Mid(A6, 1, 2)
        Case "10": E = "Sepuluh Ribu "
        Case "11": E = "Sebelas Ribu "
        Case "12": E = "Dua Belas Ribu "
        Case "13": E = "Tiga Belas Ribu "
        Case "14": E = "Empat Belas Ribu "
        Case "15": E = "Lima Belas Ribu "
        Case "16": E = "Enam Belas Ribu "
        Case "17": E = "Tujuh Belas Ribu "
        Case "18": E = "Delapan Belas Ribu "
        Case "19": E = "Sembilan Belas Ribu "
    End Select
    If Mid(A6, 1, 1) <> "1" And _
    Mid(A5, 1, 1) = "1" Then D = "Satu Ribu "
    Select Case Mid(A6, 1, 1)
        Case "2": E = "Dua Puluh "
    End Select
    Select Case Mid(A6, 1, 1) And _
    Mid(A5, 1, 1) = "0"
        Case "2": E = "Dua Puluh Ribu "
    End Select
    If Len(Angka) = 1 Then Huruf = A
    If Len(Angka) = 2 Then Huruf = B + A
    If Len(Angka) = 3 Then Huruf = C + B + A
    If Len(Angka) = 5 And _
        Mid(A5, 1, 1) = "1" Then
        D = "Seribu ": Huruf = D + C + B + A
    End If
    If Len(Angka) = 5 Then
        Huruf = D + C + B + A
    ElseIf Len(Angka) = 6 Then
        Huruf = E + D + C + B + A
    End If
    Tulis = Huruf & " Rupiah"
End Sub
