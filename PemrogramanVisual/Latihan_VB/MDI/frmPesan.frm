VERSION 5.00
Begin VB.Form frmPesan 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   840
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5310
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawWidth       =   4
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   Icon            =   "frmPesan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   840
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrPesan1 
      Left            =   4080
      Top             =   120
   End
   Begin VB.Timer tmrPesan2 
      Interval        =   50
      Left            =   4680
      Top             =   120
   End
   Begin VB.Label lblPesan 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblPesan"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   660
      MousePointer    =   11  'Hourglass
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPesan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cap$, Waktu As Byte

Property Let Pesan(Kata$) 'membuat properti Pesan
    Cap = Space(3) & UCase(Kata) & Space(3)
    Me.lblPesan.Caption = Cap
    Me.Width = lblPesan.Width + 500
    Me.Height = lblPesan.Height + 500
    Me.lblPesan.Move (Me.ScaleWidth - _
    lblPesan.Width) \ 2, (Me.ScaleHeight - _
    lblPesan.Height) \ 2: Me.Show
End Property

Sub Form_Load()
    Me.tmrPesan1.Enabled = False
    Me.tmrPesan2.Enabled = True
End Sub

Property Let Lama(Detik%) 'membuat properti Lama
    Waktu = Detik
    Me.tmrPesan1.Interval = Waktu * 1000
    Me.tmrPesan1.Enabled = True
End Property

Sub tmrPesan1_Timer()
    Unload Me
End Sub

Sub tmrPesan2_Timer()
    'animasi label berkedip
    If lblPesan.Visible = True Then
        lblPesan.Visible = False
    Else: lblPesan.Visible = True
    End If
    Me.ZOrder 'form mengambang
End Sub
