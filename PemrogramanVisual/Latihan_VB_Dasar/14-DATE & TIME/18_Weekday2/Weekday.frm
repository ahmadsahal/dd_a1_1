VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mencari hari lahir"
   ClientHeight    =   2136
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3708
   LinkTopic       =   "Form1"
   ScaleHeight     =   2136
   ScaleWidth      =   3708
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   1800
      TabIndex        =   3
      Top             =   1440
      Width           =   1692
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   372
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   1692
   End
   Begin VB.Label Label2 
      Caption         =   "Hari lahir Anda"
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal lahir Anda"
      Height          =   252
      Left            =   120
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
Private Sub Command1_Click()
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
End Sub
