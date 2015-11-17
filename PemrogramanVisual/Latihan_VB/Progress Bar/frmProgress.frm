VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Progress Bar"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdJalankan 
      Caption         =   "&Jalankan"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Timer tmrProgress 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   360
      Top             =   1080
   End
   Begin MSComctlLib.ProgressBar pgbProgress 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Min             =   1e-4
      Max             =   100
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdJalankan_Click()
    Me.tmrProgress.Enabled = True
End Sub

Private Sub tmrProgress_Timer()
    Me.pgbProgress.Value = Me.pgbProgress.Value + 1
    If Me.pgbProgress.Value = Me.pgbProgress.Max Then
        MsgBox "Percobaan selesai", vbInformation
        Me.tmrProgress.Enabled = False
        Me.pgbProgress.Value = Me.pgbProgress.Min
    End If
End Sub
