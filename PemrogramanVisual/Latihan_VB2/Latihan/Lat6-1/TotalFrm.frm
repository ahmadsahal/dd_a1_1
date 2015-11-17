VERSION 5.00
Begin VB.Form TotalFrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Total Dana"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5700
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OKBtn 
      Caption         =   "OK"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label LblDana 
      Alignment       =   1  'Right Justify
      Caption         =   "10000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Total dana"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label LblNasabah 
      Alignment       =   1  'Right Justify
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Jumlah Nasabah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "TotalFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim x As Single
Dim Nasabah As Long

'buka recordset tabel Profile & Saving
'eksekusi perintah SQL
    Cmd.CommandText = "SELECT * FROM profile"
    Set Cmd.ActiveConnection = MyDB
    'Set Rc = Cmd.Execute
    RcProfile.Open Cmd, , adOpenDynamic, adLockOptimistic
    
    Cmd.CommandText = "SELECT*FROM saving"
    RcSaving.Open Cmd, , adOpenDynamic, adLockOptimistic
    
'ambil jumlah record
RcProfile.MoveFirst
Nasabah = 1
RcProfile.MoveNext

While Not RcProfile.EOF
    Nasabah = Nasabah + 1
    RcProfile.MoveNext
Wend

LblNasabah.Caption = Nasabah


RcSaving.MoveFirst
x = RcSaving!saldo
RcSaving.MoveNext

While Not RcSaving.EOF
    x = x + RcSaving!saldo
    RcSaving.MoveNext
Wend

LblDana.Caption = Format(x, "###,###,###,###.00")

RcProfile.Close
RcSaving.Close

End Sub

Private Sub OKBtn_Click()
Unload Me
End Sub
