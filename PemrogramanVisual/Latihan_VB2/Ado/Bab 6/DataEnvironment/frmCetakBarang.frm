VERSION 5.00
Begin VB.Form frmCetakBarang 
   Caption         =   "Barang"
   ClientHeight    =   765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   765
   ScaleWidth      =   4980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCetak 
      Caption         =   "Cetak"
      Height          =   492
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1644
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "Batal"
      Height          =   492
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   1644
   End
End
Attribute VB_Name = "frmCetakBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBatal_Click()
    Unload Me
End Sub

Private Sub cmdCetak_Click()
    CetakBarang.Show
End Sub
