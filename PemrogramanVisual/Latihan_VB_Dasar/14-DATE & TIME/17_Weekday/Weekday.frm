VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
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
Private Sub Form_Activate()
    Dim KodeHari As Byte
    Dim Hari As String
    KodeHari = Weekday("14/05/03")
    Select Case KodeHari
        Case 1: Hari = "Minggu"
        Case 2: Hari = "Senin"
        Case 3: Hari = "Selasa"
        Case 4: Hari = "Rabu"
        Case 5: Hari = "Kamis"
        Case 6: Hari = "Jumat"
        Case 7: Hari = "Sabtu"
    End Select
    Print Hari
End Sub

