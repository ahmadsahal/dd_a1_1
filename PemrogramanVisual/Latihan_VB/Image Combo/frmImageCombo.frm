VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImageCombo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   2910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageCombo imcCombo 
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      ImageList       =   "imlCombo"
   End
   Begin MSComctlLib.ImageList imlCombo 
      Left            =   1080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImageCombo.frx":0000
            Key             =   "A"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImageCombo.frx":00FC
            Key             =   "B"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImageCombo.frx":0418
            Key             =   "C"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmImageCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    With Me.imcCombo
        .ComboItems.Add , "A", "Komputer", "A"
        .ComboItems.Add , "B", "Drive", "B", , 1
        .ComboItems.Add , "C", "Disk", "C", , 2
    End With
End Sub

Private Sub imcCombo_Click()
    MsgBox "Anda memilih " + Me.imcCombo.Text
End Sub
