VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form mainFrm 
   Caption         =   "Chat Client - SYA2"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   Icon            =   "mainFrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton DisconnectBtn 
      Caption         =   "Disconnect"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton ConnectBtn 
      Caption         =   "Connect"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   4560
      Width           =   6375
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4680
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "202.155.18.1"
      RemotePort      =   13480
      LocalPort       =   13480
   End
   Begin VB.TextBox Text1 
      Height          =   3615
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "mainFrm.frx":0442
      Top             =   240
      Width           =   6375
   End
   Begin VB.Label Label1 
      Caption         =   "Ketik pesan pada textbox di bawah ini, lalu tekan enter untuk mengirim"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   6375
   End
End
Attribute VB_Name = "mainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pesan As String

Private Sub ConnectBtn_Click()
Winsock1.Connect
ConnectBtn.Enabled = False
End Sub

Private Sub DisconnectBtn_Click()
Winsock1.Close
DisconnectBtn.Enabled = False
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text2.Enabled = False
DisconnectBtn.Enabled = False
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
'kalau enter ditekan
If KeyAscii = 13 Then
    If Text2.Text = "" Then Exit Sub
    Winsock1.SendData Pesan
    Pesan = "<" + Winsock1.LocalIP + "> " + Text2.Text
    Text1.Text = Text1.Text + Pesan + vbCrLf
End If
End Sub

Private Sub Winsock1_Close()
ConnectBtn.Enabled = True
DisconnectBtn.Enabled = False
Text2.Enabled = False
End Sub

Private Sub Winsock1_Connect()
ConnectBtn.Enabled = False
DisconnectBtn.Enabled = True
Text2.Enabled = True
End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData Pesan
Pesan = "<" + Winsock1.RemoteHost + "> " + Pesan
Text1.Text = Text1.Text + Pesan + vbCrLf
End Sub

Private Sub Winsock1_SendComplete()
MsgBox "Data telah terkirim ke server!", vbInformation, "Sukses!"
End Sub
