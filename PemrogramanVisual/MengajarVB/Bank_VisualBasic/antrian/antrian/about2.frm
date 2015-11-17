VERSION 5.00
Begin VB.Form aboutcompany 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tentang"
   ClientHeight    =   8130
   ClientLeft      =   4515
   ClientTop       =   1425
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "about2.frx":0000
   ScaleHeight     =   8130
   ScaleWidth      =   12135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   9960
      Top             =   2040
   End
   Begin VB.PictureBox picCredits 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5295
      Left            =   7560
      ScaleHeight     =   5295
      ScaleWidth      =   3975
      TabIndex        =   0
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Click here to exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   9600
      TabIndex        =   1
      Top             =   7440
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   8175
      Left            =   -120
      Picture         =   "about2.frx":15312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12255
   End
End
Attribute VB_Name = "aboutcompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim CreditLine()    As String
Dim CreditLeft()    As Long
Dim ColorFades(100) As Long
Dim ScrollSpeed     As Integer
Dim ColText         As Long
Dim FadeIn          As Long
Dim FadeOut         As Long

Dim cDiff1          As Long
Dim cDiff2          As Double
Dim cDiff3          As Double

Dim TotalLines      As Integer
Dim LinesOffset     As Integer
Dim Yscroll         As Long
Dim CharHeight      As Integer
Dim LinesVisible    As Integer

Private Sub Form_Load()
Dim FileO       As Integer
Dim fileName    As String
Dim tmp         As String
Dim i           As Integer

Dim Rcol1       As Long
Dim Gcol1       As Long
Dim Bcol1       As Long

Dim Rcol2       As Long
Dim Gcol2       As Long
Dim Bcol2       As Long

Dim Rfade       As Long
Dim Gfade       As Long
Dim Bfade       As Long

Dim PercentFade As Integer
Dim TimeInterval As Integer
Dim AlignText  As Integer

    PercentFade = 20
    TimeInterval = 10
    ScrollSpeed = 15
    AlignText = 2
 

LinesVisible = (picCredits.Height / picCredits.TextHeight("A")) + 1

For i = 1 To LinesVisible
    ReDim Preserve CreditLine(TotalLines) As String
    CreditLine(TotalLines) = tmp
    TotalLines = TotalLines + 1
Next

FileO = FreeFile
fileName = App.Path & "\AboutCompany.txt"
If Dir(fileName) = "" Then
        GoTo errHandler
        End If
On Error GoTo errHandler
Open fileName For Input As FileO
While Not EOF(FileO)
    Line Input #FileO, tmp
    ReDim Preserve CreditLine(TotalLines) As String
    CreditLine(TotalLines) = tmp
    TotalLines = TotalLines + 1
    Wend
Close #FileO

Me.Timer1.Interval = TimeInterval

LinesVisible = (picCredits.Height / picCredits.TextHeight("A")) + 1

CharHeight = picCredits.TextHeight("A")
If PercentFade <> 0 Then
    FadeOut = ((picCredits.Height / 100) * PercentFade) - CharHeight
    FadeIn = (picCredits.Height - FadeOut) - CharHeight - CharHeight
    Else
    FadeIn = picCredits.Height
    FadeOut = 0 - CharHeight
    End If
    
ColText = picCredits.ForeColor
cDiff1 = (picCredits.Height - (CharHeight - 10)) - FadeIn
cDiff2 = 100 / cDiff1
cDiff3 = 100 / FadeOut

ReDim CreditLeft(TotalLines - 1)
For i = 0 To TotalLines - 1
    Select Case AlignText
    Case 1
        CreditLeft(i) = 100
    Case 2
        CreditLeft(i) = (picCredits.Width - picCredits.TextWidth(CreditLine(i))) / 2
    Case 3
        CreditLeft(i) = picCredits.Width - picCredits.TextWidth(CreditLine(i)) - 100
    End Select
Next i

Rcol1 = picCredits.ForeColor Mod 256
Gcol1 = (picCredits.ForeColor And vbGreen) / 256
Bcol1 = (picCredits.ForeColor And vbBlue) / 65536
Rcol2 = picCredits.BackColor Mod 256
Gcol2 = (picCredits.BackColor And vbGreen) / 256
Bcol2 = (picCredits.BackColor And vbBlue) / 65536
For i = 0 To 100
    Rfade = Rcol2 + ((Rcol1 - Rcol2) / 100) * i: If Rfade < 0 Then Rfade = 0
    Gfade = Gcol2 + ((Gcol1 - Gcol2) / 100) * i: If Gfade < 0 Then Gfade = 0
    Bfade = Bcol2 + ((Bcol1 - Bcol2) / 100) * i: If Bfade < 0 Then Bfade = 0
    ColorFades(i) = RGB(Rfade, Gfade, Bfade)
Next


Me.Timer1.Enabled = True
Exit Sub

errHandler:
Close FileO
MsgBox "Could not load Credits", vbCritical, " Credits Demo"
End Sub


Private Sub Label1_Click(Index As Integer)
Unload Me
End Sub

Private Sub Timer1_Timer()
Dim Ycurr       As Long
Dim TextLine    As Integer
Dim ColPrct     As Long
Dim i           As Integer
picCredits.Cls
Yscroll = Yscroll - ScrollSpeed
If Yscroll < (0 - CharHeight) Then
    Yscroll = 0
    LinesOffset = LinesOffset + 1
    If LinesOffset > TotalLines - 1 Then LinesOffset = 0
    End If
picCredits.CurrentY = Yscroll
Ycurr = Yscroll
For i = 1 To LinesVisible
    If Ycurr > FadeIn And Ycurr < picCredits.Height Then
        ColPrct = cDiff2 * (cDiff1 - (Ycurr - FadeIn))
        If ColPrct < 0 Then ColPrct = 0
        If ColPrct > 100 Then ColPrct = 100
        picCredits.ForeColor = ColorFades(ColPrct)
    ElseIf Ycurr < FadeOut Then
        ColPrct = cDiff3 * Ycurr
        If ColPrct < 0 Then ColPrct = 0
        If ColPrct > 100 Then ColPrct = 100
        picCredits.ForeColor = ColorFades(ColPrct)
    Else
        picCredits.ForeColor = ColText
    End If
    TextLine = (i + LinesOffset) Mod TotalLines
    picCredits.CurrentX = CreditLeft(TextLine)
    picCredits.Print CreditLine(TextLine)
    Ycurr = Ycurr + CharHeight
Next i
End Sub

Private Sub Form_Click()
Unload Me
End Sub






