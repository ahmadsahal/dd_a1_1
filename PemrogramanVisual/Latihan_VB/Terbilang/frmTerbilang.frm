VERSION 5.00
Begin VB.Form frmTerbilang 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Terbilang"
   ClientHeight    =   3525
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5760
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTerbilang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3195
      ScaleWidth      =   5475
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.TextBox txt1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H00FF0000&
         Height          =   390
         Left            =   120
         MaxLength       =   14
         TabIndex        =   4
         Top             =   225
         Width           =   5175
      End
      Begin VB.CommandButton cmd1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Terbilang"
         Height          =   495
         Left            =   3600
         MaskColor       =   &H00FF00FF&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2640
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox txt2 
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H00FF0000&
         Height          =   1815
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   5175
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Clear"
         Height          =   495
         Left            =   120
         MaskColor       =   &H00FF00FF&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2640
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmTerbilang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Terbilang()
    Dim Angka, Huruf
    '*********************************************
    Dim A, B, C, D, E, F, G, H, I, J, K, L, M
    Dim A1, A2, A3, A5, A6, A7
    Dim A8, A9, A10, A11, A12, A13, A14, A15
    '*********************************************
    Angka = txt1.Text
    '*********************************************
    A1 = Right(Angka, 1): A5 = Right(Angka, 5)
    A2 = Right(Angka, 2): A6 = Right(Angka, 6)
    A3 = Right(Angka, 3): A7 = Right(Angka, 7)

    A9 = Right(Angka, 9): A13 = Right(Angka, 13)
    A10 = Right(Angka, 10): A14 = Right(Angka, 14)
    A11 = Right(Angka, 11):  A15 = Right(Angka, 15)
    '**********************************************
        Select Case A1
                Case "1": A = "Satu "
                Case "2": A = "Dua "
                Case "3": A = "Tiga "
                Case "4": A = "Empat "
                Case "5": A = "Lima "
                Case "6": A = "Enam "
                Case "7": A = "Tujuh "
                Case "8": A = "Delapan "
                Case "9": A = "Sembilan "
        End Select
    '**********************************************
        Select Case A2
                Case "10": A = "Sepuluh "
                Case "11": A = "Sebelas "
                Case "12": A = "Dua Belas "
                Case "13": A = "Tiga Belas "
                Case "14": A = "Empat Belas "
                Case "15": A = "Lima Belas "
                Case "16": A = "Enam Belas "
                Case "17": A = "Tujuh Belas "
                Case "18": A = "Delapan Belas "
                Case "19": A = "Sembilan Belas "
        End Select
    '**********************************************
        Select Case Mid(A2, 1, 1)
            Case "2": B = "Dua Puluh "
            Case "3": B = "Tiga Puluh "
            Case "4": B = "Empat Puluh "
            Case "5": B = "Lima Puluh "
            Case "6": B = "Enam Puluh "
            Case "7": B = "Tujuh Puluh "
            Case "8": B = "Delapan Puluh "
            Case "9": B = "Sembilan Puluh "
        End Select
    '***********************************************
        Select Case Mid(A3, 1, 1)
            Case "1": C = "Seratus "
            Case "2": C = "Dua Ratus "
            Case "3": C = "Tiga Ratus "
            Case "4": C = "Empat Ratus "
            Case "5": C = "Lima Ratus "
            Case "6": C = "Enam Ratus "
            Case "7": C = "Tujuh Ratus "
            Case "8": C = "Delapan Ratus "
            Case "9": C = "Sembilan Ratus "
        End Select
    '************************************************
        Select Case Mid(A5, 1, 1) _
        And Mid(A6, 1, 1) <> "1"
            Case "2": D = "Dua Ribu "
            Case "3": D = "Tiga Ribu "
            Case "4": D = "Empat Ribu "
            Case "5": D = "Lima Ribu "
            Case "6": D = "Enam Ribu "
            Case "7": D = "Tujuh Ribu "
            Case "8": D = "Delapan Ribu "
            Case "9": D = "Sembilan Ribu "
        End Select
        
    '**********************************************
        Select Case Mid(A6, 1, 2)
            Case "10": E = "Sepuluh Ribu "
            Case "11": E = "Sebelas Ribu "
            Case "12": E = "Dua Belas Ribu "
            Case "13": E = "Tiga Belas Ribu "
            Case "14": E = "Empat Belas Ribu "
            Case "15": E = "Lima Belas Ribu "
            Case "16": E = "Enam Belas Ribu "
            Case "17": E = "Tujuh Belas Ribu "
            Case "18": E = "Delapan Belas Ribu "
            Case "19": E = "Sembilan Belas Ribu "
        End Select
    '*************************************************
        If Mid(A6, 1, 1) <> "1" And _
            Mid(A5, 1, 1) = "1" Then
            D = "Satu Ribu "
        End If
    '*************************************************
        Select Case Mid(A6, 1, 1)
            Case "2": E = "Dua Puluh "
            Case "3": E = "Tiga Puluh "
            Case "4": E = "Empat Puluh "
            Case "5": E = "Lima Puluh "
            Case "6": E = "Enam Puluh "
            Case "7": E = "Tujuh Puluh "
            Case "8": E = "Delapan Puluh "
            Case "9": E = "Sembilan Puluh "
        End Select
    '************************************************
     Select Case Mid(A6, 1, 1) And Mid(A5, 1, 1) = "0"
            Case "2": E = "Dua Puluh Ribu "
            Case "3": E = "Tiga Puluh Ribu "
            Case "4": E = "Empat Puluh Ribu "
            Case "5": E = "Lima Puluh Ribu "
            Case "6": E = "Enam Puluh Ribu "
            Case "7": E = "Tujuh Puluh Ribu "
            Case "8": E = "Delapan Puluh Ribu "
            Case "9": E = "Sembilan Puluh Ribu "
    End Select
     '*************************************************
        Select Case Mid(A7, 1, 1)
            Case "1": F = "Seratus "
            Case "2": F = "Dua Ratus "
            Case "3": F = "Tiga Ratus "
            Case "4": F = "Empat Ratus "
            Case "5": F = "Lima Ratus "
            Case "6": F = "Enam Ratus "
            Case "7": F = "Tujuh Ratus "
            Case "8": F = "Delapan Ratus  "
            Case "9": F = "Sembilan Ratus "
        End Select
    '*************************************************
        Select Case Mid(A7, 1, 1) And _
        Mid(A6, 1, 2) = "00"
            Case "1": F = "Seratus Ribu "
            Case "2": F = "Dua Ratus Ribu "
            Case "3": F = "Tiga Ratus Ribu "
            Case "4": F = "Empat Ratus Ribu "
            Case "5": F = "Lima Ratus Ribu "
            Case "6": F = "Enam Ratus Ribu "
            Case "7": F = "Tujuh Ratus Ribu "
            Case "8": F = "Delapan Ratus Ribu "
            Case "9": F = "Sembilan Ratus Ribu "
        End Select
    '**************************************************
        Select Case Mid(A9, 1, 1)           '21,111,123
            Case "1": G = "Satu Juta "
            Case "2": G = "Dua Juta "
            Case "3": G = "Tiga Juta "
            Case "4": G = "Empat Juta "
            Case "5": G = "Lima Juta "
            Case "6": G = "Enam Juta "
            Case "7": G = "Tujuh Juta "
            Case "8": G = "Delapan Juta "
            Case "9": G = "Sembilan Juta "
        End Select
     '*************************************************
        Select Case Mid(A10, 1, 2)
            Case "10": G = "Sepuluh Juta "
            Case "11": G = "Sebelas Juta "
            Case "12": G = "Dua Belas Juta "
            Case "13": G = "Tiga Belas Juta "
            Case "14": G = "Empat Belas Juta "
            Case "15": G = "Lima Belas Juta "
            Case "16": G = "Enam Belas Juta "
            Case "17": G = "Tujuh Belas Juta "
            Case "18": G = "Delapan Belas Juta "
            Case "19": G = "Sembilan Belas Juta "
        End Select
    '************************************************
        Select Case Mid(A10, 1, 1)
            Case "2": H = "Dua Puluh "
            Case "3": H = "Tiga Puluh "
            Case "4": H = "Empat Puluh "
            Case "5": H = "Lima Puluh "
            Case "6": H = "Enam Puluh "
            Case "7": H = "Tujuh Puluh "
            Case "8": H = "Delapan Puluh "
            Case "9": H = "Sembilan Puluh "
        End Select
      '**************************************************
        Select Case Mid(A10, 1, 1) And _
        Mid(A9, 1, 1) = "0"
            Case "2": H = "Dua Puluh Juta "
            Case "3": H = "Tiga Puluh Juta "
            Case "4": H = "Empat Puluh Juta "
            Case "5": H = "Lima Puluh Juta "
            Case "6": H = "Enam Puluh Juta "
            Case "7": H = "Tujuh Puluh Juta "
            Case "8": H = "Delapan Puluh Juta "
            Case "9": H = "Sembilan Puluh Juta "
        End Select
       '**************************************************
        Select Case Mid(A11, 1, 1)
            Case "1": I = "Seratus "
            Case "2": I = "Dua Ratus "
            Case "3": I = "Tiga Ratus "
            Case "4": I = "Empat Ratus "
            Case "5": I = "Lima Ratus "
            Case "6": I = "Enam Ratus "
            Case "7": I = "Tujuh Ratus "
            Case "8": I = "Delapan Ratus "
            Case "9": I = "Sembilan Ratus "
        End Select
        '**************************************************
        Select Case Mid(A11, 1, 1) And Mid _
        (A10, 1, 2) = "00"
            Case "1": I = "Seratus Juta "
            Case "2": I = "Dua Ratus Juta "
            Case "3": I = "Tiga Ratus Juta "
            Case "4": I = "Empat Ratus Juta "
            Case "5": I = "Lima Ratus Juta "
            Case "6": I = "Enam Ratus Juta "
            Case "7": I = "Tujuh Ratus Juta "
            Case "8": I = "Delapan Ratus Juta "
            Case "9": I = "Sembilan Ratus Juta "
        End Select
     '**************************************************
        Select Case Mid(A13, 1, 1) And _
        Mid(A14, 1, 1) <> "1"
            Case "1": J = "Satu Milyar "
            Case "2": J = "Dua Milyar "
            Case "3": J = "Tiga Milyar "
            Case "4": J = "Empat Milyar "
            Case "5": J = "Lima Milyar "
            Case "6": J = "Enam Milyar "
            Case "7": J = "Tujuh Milyar "
            Case "8": J = "Delapan Milyar "
            Case "9": J = "Sembilan Milyar "
        End Select
    '*************************************************
            Select Case Mid(A14, 1, 1)  '22,345,678,890
            Case "2": K = "Dua Puluh "
            Case "3": K = "Tiga Puluh "
            Case "4": K = "Empat Puluh "
            Case "5": K = "Lima Puluh "
            Case "6": K = "Enam Puluh "
            Case "7": K = "Tujuh Puluh "
            Case "8": K = "Delapan Puluh "
            Case "9": K = "Sembilan Puluh "
        End Select
       '**************************************************
        Select Case Mid(A14, 1, 2)
            Case "10": K = "Sepuluh Milyar "
            Case "11": K = "Sebelas Milyar "
            Case "12": K = "Dua Belas Milyar "
            Case "13": K = "Tiga Belas Milyar "
            Case "14": K = "Empat Belas Milyar "
            Case "15": K = "Lima Belas Milyar "
            Case "16": K = "Enam Belas Milyar "
            Case "17": K = "Tujuh Belas Milyar "
            Case "18": K = "Delapan Belas Milyar "
            Case "19": K = "Sembilan Belas Milyar "
        End Select
    '**************************************************

    If Len(Angka) = 1 Then
        Huruf = A
    ElseIf Len(Angka) = 2 Then
        Huruf = B + A
    ElseIf Len(Angka) = 3 Then
        Huruf = C + B + A
    ElseIf Len(Angka) = 5 And _
    Mid(A5, 1, 1) = "1" Then
        D = "Seribu"
        Huruf = D + C + B + A
    ElseIf Len(Angka) = 5 Then
        Huruf = D + C + B + A
    ElseIf Len(Angka) = 6 Then
        Huruf = E + D + C + B + A  '50,000
    ElseIf Len(Angka) = 7 Then
        Huruf = F + E + D + C + B + A
    ElseIf Len(Angka) = 9 Then
        Huruf = G + F + E + D + C + B + A
    ElseIf Len(Angka) = 10 Then
        Huruf = H + G + F + E + D + C + B + A
    ElseIf Len(Angka) = 11 Then
        Huruf = I + H + G + F + E + D + C + B + A
    ElseIf Len(Angka) = 13 Then
        Huruf = J + I + H + G + F + E + D + C + B + A
    ElseIf Len(Angka) = 14 Then '23,123,123,123
        Huruf = K + J + I + H + G + F + E + D + C + B + A
    End If
    txt2.Text = Huruf & "Rupiah." & vbCrLf
End Sub

Private Sub cmd1_Click()
On Error Resume Next
    Terbilang
    If Err Then
        MsgBox "Masukan angka yang ingin di konversikan"
        txt1.SetFocus
    End If
End Sub

Private Sub Command1_Click()
    txt1.Text = ""
    txt2.Text = ""
    txt1.SetFocus
End Sub

Private Sub Form_Activate()
    txt1.SetFocus
End Sub

Private Sub txt1_Change()
    Me.txt1.SelStart = Len(Me.txt1.Text)
    Me.txt1.Text = Format(txt1, _
    "###,###,###")
End Sub

Private Sub txt1_KeyPress( _
KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack Or _
        KeyAscii >= Asc("0") And _
        KeyAscii <= Asc("9") Or _
        KeyAscii = 13) Then
    KeyAscii = 0
    End If
End Sub
