VERSION 5.00
Begin VB.Form frmPesan 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   600
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   8355
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
   ScaleHeight     =   600
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmr1 
      Interval        =   1
      Left            =   0
      Top             =   600
   End
   Begin VB.Timer tmr2 
      Left            =   1200
      Top             =   600
   End
   Begin VB.Timer tmr3 
      Interval        =   50
      Left            =   1920
      Top             =   600
   End
   Begin VB.Label lblPesan 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   270
      Left            =   1800
      MousePointer    =   11  'Hourglass
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmPesan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit: Public cap, Lama

Property Let Lable(C$): cap = Space(3) & UCase(C$) & Space(3)
tmr1.Enabled = True: lblPesan.Caption = cap: Me.Width = lblPesan.Width + 150
Me.Height = lblPesan.Height + 150: Me.lblPesan.Move (Me.ScaleWidth - _
lblPesan.Width) \ 2, (Me.ScaleHeight - lblPesan.Height) \ 2: Me.Show: End Property

Sub Form_Load():    Me.tmr2.Enabled = False: Me.tmr1.Enabled = True
Me.tmr3.Enabled = True: End Sub

Property Let Waktu(L): Lama = L: tmr2.Interval = Lama * 1000: tmr2.Enabled = True
End Property

Sub tmr2_Timer(): Unload Me: End Sub

Sub tmr3_Timer()
If lblPesan.Visible = True Then
lblPesan.Visible = False
Else: lblPesan.Visible = True: End If: Me.ZOrder: End Sub

