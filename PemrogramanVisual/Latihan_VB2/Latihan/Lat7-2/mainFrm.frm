VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form mainFrm 
   Caption         =   "Chat Server - IWAN"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9525
   Icon            =   "mainFrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   9525
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   4350
      ItemData        =   "mainFrm.frx":0442
      Left            =   6840
      List            =   "mainFrm.frx":0449
      TabIndex        =   3
      Top             =   600
      Width           =   2415
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
      LocalPort       =   13480
   End
   Begin VB.TextBox Text1 
      Height          =   3615
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "mainFrm.frx":045B
      Top             =   240
      Width           =   6375
   End
   Begin VB.Label Label2 
      Caption         =   "IP Address Client"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   120
      Width           =   2415
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
   Begin VB.Menu mnuPopup 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu MnuPutus 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu MnuHost 
         Caption         =   "HostName"
      End
   End
End
Attribute VB_Name = "mainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pesan As String
Dim Host As String
Dim Idx As Integer



Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Winsock1.Listen
'List1.Clear
End Sub



Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If List1.ListCount < 1 Then Exit Sub
Host = List1.Text
Idx = List1.ListIndex
'kalau tombol kanan mouse diklik...
If Button = 2 Then
    PopupMenu mnuPopup
End If

End Sub

Private Sub MnuHost_Click()
MsgBox Winsock1.RemoteHostIP + ", " + Winsock1.RemoteHost + " connected from port " + Str(Winsock1.RemotePort), vbInformation, "Info host " + Winsock1.RemoteHost
End Sub

Private Sub MnuPutus_Click()
Winsock1.Close
Winsock1.Listen
List1.RemoveItem Idx
List1.Refresh
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
List1.Clear
Winsock1.Close
Winsock1.Listen
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Accept requestID
List1.AddItem Winsock1.RemoteHost
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData Pesan
Pesan = "<" + Winsock1.RemoteHost + "> " + Pesan
Text1.Text = Text1.Text + Pesan + vbCrLf
End Sub

Private Sub Winsock1_SendComplete()
MsgBox "Data telah terkirim ke server!", vbInformation, "Sukses!"
End Sub
