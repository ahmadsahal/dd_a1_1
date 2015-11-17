VERSION 5.00
Begin VB.Form frmCHR 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3780
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCHR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   8000
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Untuk mengetahui kepribadian Anda sesungguhnya, silakan hubungi viansastra@ telkom.net"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   810
      Left            =   435
      TabIndex        =   2
      Top             =   2760
      Width           =   6660
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCHR.frx":000C
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   1350
      Left            =   375
      TabIndex        =   1
      Top             =   360
      Width           =   6660
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCHR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ILEGAL"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   405
      Left            =   3000
      TabIndex        =   0
      Top             =   2040
      Width           =   1260
   End
End
Attribute VB_Name = "frmCHR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
Me.Label1.Caption = "Saudara " + UCase(NAMA) + " di " + Alamat + ", ternyata setelah membandingkan" + _
" nilai dari seluruh kelompok pertanyaan, maka Saudara " + NAMA + " termasuk dalam golongan:"
If TotalKolom1 > TotalKolom2 And TotalKolom1 > TotalKolom3 And TotalKolom1 > TotalKolom4 Then
Me.lblCHR.Caption = "SANGUINIS": End If
If TotalKolom2 > TotalKolom1 And TotalKolom2 > TotalKolom3 And TotalKolom2 > TotalKolom4 Then
Me.lblCHR.Caption = "KOLERIS": End If
If TotalKolom3 > TotalKolom1 And TotalKolom3 > TotalKolom2 And TotalKolom3 > TotalKolom4 Then
Me.lblCHR.Caption = "MELANKOLIS": End If
If TotalKolom4 > TotalKolom2 And TotalKolom4 > TotalKolom3 And TotalKolom4 > TotalKolom1 Then
Me.lblCHR.Caption = "PLEGMATIS": End If
End Sub
Private Sub Form_Resize()
'Me.lblCHR.Left = (Me.Width - Me.lblCHR.Width) / 2
'Me.lblCHR.Top = (Me.Height - Me.lblCHR.Height) / 2
'Me.Label1.Left = (Me.Width - Me.Label1.Width) / 2
End Sub
Private Sub Timer1_Timer()
If Me.lblCHR.Visible = True Then
Me.lblCHR.Visible = False
Else: Me.lblCHR.Visible = True: End If
End Sub

Private Sub Timer2_Timer()
End
End Sub
