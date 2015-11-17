VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
'Trim: Untuk menghilangkan spasi
'dari kiri dan kanan teks
    a = "  Tom   "
    B = "  Jerry "
    Print "Tanpa Trim() Panjang text =" & Len(a) & " character Textnya=>" & Trim(a) & " Setelah = " & Len(Trim(a))
    Print "Tanpa Trim() Panjang text =" & Len(B) & " character Textnya=>" & Trim(B) & " Setelah = " & Len(Trim(B))
    Print Trim(a) + Trim(B) 'TomJerry
End Sub

 
