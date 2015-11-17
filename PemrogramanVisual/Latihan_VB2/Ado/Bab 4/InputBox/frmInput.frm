VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1065
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Input Ibukota Indonesia"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim jawaban As String
    Dim defaul As String
    Dim Pesan As String
    Dim JudulWindow As String
    
    Pesan = "Masukkan Nama Ibukota Indonesia"
    JudulWindow = "Demo InputBox"
    
    defaul = "A"
    
    'Dapatkan input dari pemakai
    jawaban = InputBox(Pesan, JudulWindow, defaul)

    Select Case jawaban
        Case "JAKARTA"
            Text1.Text = jawaban & " Ibukotanya Indonesia"
        Case "Jakarta"
            Text1.Text = jawaban & " Ibukotanya Indonesia"
        Case "jakarta"
            Text1.Text = jawaban & " Ibukotanya Indonesia"
        Case ""
            MsgBox "Proses Dibatalkan"
            Text1.Text = ""
        Case Else
            Text1.Text = jawaban & "Tidak terdapat pada database"
    End Select
End Sub

