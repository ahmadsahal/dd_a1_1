VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTreeView 
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imlTreeView 
      Left            =   2040
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeView.frx":0000
            Key             =   "A"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeView.frx":031C
            Key             =   "B"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeView.frx":0638
            Key             =   "C"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeView.frx":0F14
            Key             =   "D"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trvTreeView 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5106
      _Version        =   393217
      Style           =   7
      ImageList       =   "imlTreeView"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim i
    With trvTreeView.Nodes
        .Add , 1, "A", "Master", "A"
        .Add 1, 4, "B", "Adodc", "B"
        .Add 1, 4, "C", "ListView", "B"
        .Add 1, 4, "D", "File", "B"
        .Add 1, 4, "E", "Multimedia", "B"
        .Add 5, 4, "F", "Objek", "C"
        .Add 5, 4, "G", "Kode", "C"
        .Add , 3, "H", "About", "D"
        .Add 8, 4, "I", "Penulis", "B"
    End With
End Sub


Private Sub trvTreeView_NodeClick(ByVal _
Node As MSComctlLib.Node)
If Not (Node.Key = "A") Then
    Select Case Node.Key
        Case "B"
            MsgBox "Anda memilih cabang " _
            + Node.Text
        Case "C"
            MsgBox "Anda memilih cabang " _
            + Node.Text
        Case Else
            MsgBox "Anda memilih cabang " _
            + Node.Text
    End Select
End If
End Sub
