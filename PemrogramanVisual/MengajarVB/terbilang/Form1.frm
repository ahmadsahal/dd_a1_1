VERSION 5.00
Begin VB.Form FormTerbilang 
   Caption         =   "Form1"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox ListTerbilang 
      Height          =   1425
      Left            =   3960
      TabIndex        =   6
      Top             =   5400
      Width           =   4335
   End
   Begin VB.ListBox Langka 
      Height          =   1620
      Left            =   360
      TabIndex        =   5
      Top             =   5400
      Width           =   3135
   End
   Begin VB.TextBox txtTerbilang 
      Height          =   4455
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   7455
   End
   Begin VB.TextBox txtDuit 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdProses 
      Caption         =   "Proses"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Terbilang"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Rupiah"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "FormTerbilang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function duitRp(Rupiah As Double)
   Dim Angka(0 To 9) As String
   Dim s4, s3, s2, s1, kon, bulat, Desimal As String
   Dim i, C, Bagian As Integer
   Dim Check, d1, d2, d3 As Double
   s4 = " Ratus "
   s3 = " Ribu "
   s2 = " Juta "
   s1 = " Milyard "
   Angka(1) = " satu"
   Angka(2) = " dua"
   Angka(3) = " tiga"
   Angka(4) = " empat"
   Angka(5) = " lima"
   Angka(6) = " enam"
   Angka(7) = " tujuh"
   Angka(8) = " delapan"
   Angka(9) = " sembilan"
   Angka(0) = " nol"
   kon = " "
   bulat = Mid$(Format(Rupiah, "000000000000.00"), 1, 12)
   Desimal = Mid$(Format(Rupiah, "000000000000.00"), 14, 2)
   i = 1
   Bagian = 1
   C = 1
   
   Do While i <= 4
      Sudah = False
      Check = Mid$(bulat, (i - 1) * 3 + 1)
      d1 = Val(Mid$(Check, 1, 1))
      d2 = Val(Mid$(Check, 2, 1))
      d3 = Val(Mid$(Check, 3, 1))
      If Not (d1 + d2 + d3) = 0 Then
         If d1 = 1 Then
            kon = kon + "Seratus "
         Else
            If Not (d1 = 0) Then
               kon = kon + Trim(Angka(d1)) + " Ratus "
            End If
         End If
         If d2 = 1 Then
            If d3 = 0 Then
               kon = kon + "Sepuluh "
            ElseIf d3 = 1 Then
               kon = kon + " Se"
            Else
               kon = kon
            End If
         
            If Not (d3 = 0) Then
               Sudah = True
               If d3 = 1 Then
                  kon = kon + "belas "
               Else
                  kon = kon + Trim(Angka(d3)) + " Belas "
               End If
            End If
         Else
            If Not (d2 = 0) Then
               kon = kon + Trim(Angka(d2)) + " Puluh "
            End If
         End If
         If d3 = 1 And Not (Sudah) Then
             If i = 3 Then
                If d1 + d2 = 0 Then
                   kon = kon + "Se"
                Else
                   kon = kon + "Satu "
                End If
             Else
                kon = kon + "Satu "
             End If
            
         Else
            If Not (d3 = 0) And Not (Sudah) Then
               kon = kon + Trim(Angka(d3))
            End If
         End If
         
         If Not (i = 4) Then
            Select Case i
                    Case 1
                       kon = kon + s1
                    Case 2
                       kon = kon + s2
                    Case 3
                       kon = kon + s3
            End Select
            
            If i = 3 And d1 + d2 = 0 And d3 = 1 Then
               kon = "Seribu "
            End If
         End If
      End If
      i = i + 1
   Loop
   If Val(Desimal) <> 0 Then
      kon = kon + " Point " '+ duitRp2(Val(desimal))
      For i = 1 To 2
          kon = kon + Angka(Val(Mid$(Desimal, i, 1)))
      Next i
   End If
   duitRp = UCase(kon) + " RUPIAH"
End Function

Private Sub cmdProses_Click()
Dim i As Integer
Dim Bil As Double
'txtTerbilang.Text = duitRp(txtDuit)
For i = 0 To Langka.ListCount - 1
'Bil = duitRp((Langka.List(i)))
ListTerbilang.AddItem duitRp((Langka.List(i)))
txtTerbilang.Text = txtTerbilang.Text + Chr(13) + duitRp((Langka.List(i)))


'Langka.ItemData(i).Text

Next i


End Sub

Private Sub Form_Load()
Langka.AddItem 375000
Langka.AddItem 375000
Langka.AddItem 375000
Langka.AddItem 375000
Langka.AddItem 375000
Langka.AddItem 375000
Langka.AddItem 0
Langka.AddItem 1015200
Langka.AddItem 0
Langka.AddItem 450000
Langka.AddItem 450000
Langka.AddItem 450000
Langka.AddItem 450000
Langka.AddItem 0
Langka.AddItem 600000
Langka.AddItem 600000
Langka.AddItem 600000
Langka.AddItem 600000
Langka.AddItem 600000
Langka.AddItem 600000







End Sub
