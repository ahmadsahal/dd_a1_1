VERSION 5.00
Begin {90290CCD-F27D-11D0-8031-00C04FB6C701} DHTMLPage1 
   ClientHeight    =   5712
   ClientLeft      =   1812
   ClientTop       =   1548
   ClientWidth     =   6276
   _ExtentX        =   11070
   _ExtentY        =   10075
   SourceFile      =   ""
   BuildFile       =   ""
   BuildMode       =   0
   TypeLibCookie   =   977
   AsyncLoad       =   0   'False
   id              =   "DHTMLPage1"
   ShowBorder      =   -1  'True
   ShowDetail      =   -1  'True
   AbsPos          =   0   'False
   HTMLDocument    =   "Contoh.dsx":0000
End
Attribute VB_Name = "DHTMLPage1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Function Button1_onclick() As Boolean
    'MENCARI HARI LAHIR
    Dim KodeHari As Byte
    Dim Hari As String
    KodeHari = Weekday(TextField1.Value)
    Select Case KodeHari
        Case 1: Hari = "Minggu"
        Case 2: Hari = "Senin"
        Case 3: Hari = "Selasa"
        Case 4: Hari = "Rabu"
        Case 5: Hari = "Kamis"
        Case 6: Hari = "Jumat"
        Case 7: Hari = "Sabtu"
    End Select
    TextField2.innerText = Hari
    
    'MENCARI BINTANG
    Dim Tg, Bl As Byte
    Dim Btg As String
    Tg = Day(TextField1.Value): Bl = Month(TextField1.Value)
    If (Tg >= 21 And Bl = 3) Or (Tg <= 20 And Bl = 4) Then Btg = "ARIES"
    If (Tg >= 21 And Bl = 4) Or (Tg <= 21 And Bl = 5) Then Btg = "TAURUS"
    If (Tg >= 22 And Bl = 5) Or (Tg <= 21 And Bl = 6) Then Btg = "GEMINI"
    If (Tg >= 22 And Bl = 6) Or (Tg <= 22 And Bl = 7) Then Btg = "CANCER"
    If (Tg >= 23 And Bl = 7) Or (Tg <= 22 And Bl = 8) Then Btg = "LEO"
    If (Tg >= 23 And Bl = 8) Or (Tg <= 22 And Bl = 9) Then Btg = "VIRGO"
    If (Tg >= 23 And Bl = 9) Or (Tg <= 22 And Bl = 10) Then Btg = "LIBRA"
    If (Tg >= 23 And Bl = 10) Or (Tg <= 21 And Bl = 11) Then Btg = "SCORPIO"
    If (Tg >= 22 And Bl = 11) Or (Tg <= 21 And Bl = 12) Then Btg = "SAGITARIUS"
    If (Tg >= 22 And Bl = 12) Or (Tg <= 20 And Bl = 1) Then Btg = "CAPRICORN"
    If (Tg >= 21 And Bl = 1) Or (Tg <= 19 And Bl = 2) Then Btg = "AQUARIUS"
    If (Tg >= 20 And Bl = 2) Or (Tg <= 20 And Bl = 3) Then Btg = "PISCES"
    TextField3.innerText = Btg
End Function
