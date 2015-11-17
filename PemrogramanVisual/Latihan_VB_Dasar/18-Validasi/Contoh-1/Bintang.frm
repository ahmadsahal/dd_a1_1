VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mencari hari lahir & bintang"
   ClientHeight    =   2568
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3984
   LinkTopic       =   "Form1"
   ScaleHeight     =   2568
   ScaleWidth      =   3984
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   288
      Left            =   1920
      TabIndex        =   5
      Top             =   1920
      Width           =   1692
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   1920
      TabIndex        =   3
      Top             =   1440
      Width           =   1692
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   372
      Left            =   2640
      TabIndex        =   1
      Top             =   840
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   1692
   End
   Begin VB.Label Label3 
      Caption         =   "Bintang Anda"
      Height          =   252
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "Hari lahir Anda"
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal lahir Anda"
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1452
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Command1.Enabled = False
End Sub

Private Sub Text1_Change()
    If IsDate(Text1) Then
        Command1.Enabled = True
    Else
        Command1.Enabled = False
    End If
End Sub

Private Sub Command1_Click()
    'MENCARI HARI LAHIR
    Dim KodeHari As Byte
    Dim Hari As String
    KodeHari = Weekday(Text1)
    Select Case KodeHari
        Case 1: Hari = "Minggu"
        Case 2: Hari = "Senin"
        Case 3: Hari = "Selasa"
        Case 4: Hari = "Rabu"
        Case 5: Hari = "Kamis"
        Case 6: Hari = "Jumat"
        Case 7: Hari = "Sabtu"
    End Select
    Text2 = Hari
    
    'MENCARI BINTANG
    Dim Tg, Bl As Byte
    Dim Btg As String
    Tg = Day(Text1): Bl = Month(Text1)
    If (Tg >= 21 And Bl = 3) Or (Tg <= 20 And Bl = 4) Then Btg = "ARIES"
    If (Tg >= 21 And Bl = 4) Or (Tg <= 21 And Bl = 5) Then Btg = "TAURUS"
    If (Tg >= 22 And Bl = 5) Or (Tg <= 21 And Bl = 6) Then Btg = "GEMINI"
    If (Tg >= 22 And Bl = 6) Or (Tg <= 22 And Bl = 7) Then Btg = "CANCER"
    If (Tg >= 23 And Bl = 7) Or (Tg <= 22 And Bl = 8) Then Btg = "LEO"
    If (Tg >= 23 And Bl = 8) Or (Tg <= 22 And Bl = 9) Then Btg = "VIRGO"
    If (Tg >= 23 And Bl = 9) Or (Tg <= 22 And Bl = 10) Then Btg = "LIBRA"
    If (Tg >= 23 And Bl = 10) Or (Tg <= 21 And Bl = 11) Then Btg = "SCORPIO"
    If (Tg >= 22 And Bl = 11) Or (Tg <= 21 And Bl = 12) Then Btg = "SAGITARIUS"
    If (Tg >= 22 And Bl = 12) Or (Tg <= 20 And Bl = 1) Then Btg = "CAPRICORN"
    If (Tg >= 21 And Bl = 1) Or (Tg <= 19 And Bl = 2) Then Btg = "AQUARIUS"
    If (Tg >= 20 And Bl = 2) Or (Tg <= 20 And Bl = 3) Then Btg = "PISCES"
    Text3 = Btg
End Sub

