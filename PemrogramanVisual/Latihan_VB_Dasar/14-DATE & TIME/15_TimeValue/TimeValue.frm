VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   10.2
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TimeValue: Untuk mengkonversi
'jam string ke dalam jam

Private Sub Form_Activate()
    Dim MulaiKerja, SelesaiKerja As Variant
    MulaiKerja = TimeValue("11:30")
    SelesaiKerja = TimeValue("13:30")
    Print "Mulai kerja : "; MulaiKerja
    Print "Selesai kerja : "; SelesaiKerja
    Print "Lama kerja = "; (SelesaiKerja - MulaiKerja) * 24; "jam"
End Sub

