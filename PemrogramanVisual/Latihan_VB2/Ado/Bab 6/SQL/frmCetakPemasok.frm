VERSION 5.00
Begin VB.Form frmCetakPemasok 
   Caption         =   "Cetak Pemasok"
   ClientHeight    =   765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
   LinkTopic       =   "Form3"
   ScaleHeight     =   765
   ScaleWidth      =   4980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCetak 
      Caption         =   "Cetak"
      Height          =   492
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   1644
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "Batal"
      Height          =   492
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   1644
   End
End
Attribute VB_Name = "frmCetakPemasok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim conAVB As ADODB.Connection
Dim rsPemasok As ADODB.Recordset
Dim P As Printer

Private Sub cmdCetak_Click()
    CetakPemasok
    Printer.EndDoc
End Sub
Private Sub CetakPemasok()
    'membuat connection
    Dim strsql As String
    Set rsPemasok = New ADODB.Recordset
    strsql = "Select * from Pemasok "
    rsPemasok.Open strsql, conAVB, adOpenDynamic, adLockOptimistic, adCmdText
    'mencetak urut kode Pemasok
    Dim Nomor As Integer
    Dim Halaman As Integer
    Dim MBaris As Integer
    
    Dim MGrs As String
    On Error GoTo SalahCetak
    With rsPemasok
     'bawa pointer ke awal record yang dicetak
     .MoveFirst
     'tentukan font
     Printer.Font = "Times New Roman"
     
     'bawa head Printer ke awal halaman
     Printer.CurrentX = 0
     Printer.CurrentY = 0
     'mulai pengulangan
     Nomor = 0
     Halaman = 0
     Do While Not rsPemasok.EOF
        'cetak judul tabel
        Halaman = Halaman + 1
        Printer.Print "DAFTAR Pemasok";
        Printer.Print Tab(60); "Hal :"; Halaman
        MGrs = String$(70, "-")
        Printer.Print MGrs
        Printer.Print Tab(2); "No.";
        Printer.Print Tab(8); "Kode";
        Printer.Print Tab(20); "Nama";
        Printer.Print Tab(30); "Alamat";
        Printer.Print Tab(50); "No"
        Printer.Print Tab(2);
        Printer.Print Tab(8); "Pemasok";
        Printer.Print Tab(20); "Pemasok";
        Printer.Print Tab(30); "Pemasok";
        Printer.Print Tab(50); "Telepon"
        
        Printer.Print MGrs
        MBaris = 0
        'mulai pengulangan cetak isi tabel
        Do While Not rsPemasok.EOF And MBaris <= 22
        Nomor = Nomor + 1
           
            Printer.Print Tab(2); Nomor;
            Printer.Print Tab(8); ![Kode Pemasok];
            Printer.Print Tab(20); ![Nama Pemasok];
            Printer.Print Tab(30); ![ALamat Pemasok];
            Printer.Print Tab(50); ![no telepon]
        MBaris = MBaris + 1
           .MoveNext
           If .EOF Then
              Exit Do
           End If
        Loop
       
       Printer.NewPage
       If .EOF Then
          Exit Do
       End If
     Loop
     Printer.EndDoc
'
    End With
    On Error GoTo 0
    CmdBatal_Click
    Exit Sub
SalahCetak:
    Beep
    Dim x As Byte
    x = MsgBox("Printer ERROR!" & Chr(13) & "Betulkan printer, lalu klik OK", vbOKCancel)
    If x = 0 Then
       Resume
    Else
       Printer.KillDoc
       CmdBatal_Click
    End If
End Sub
Private Sub CmdBatal_Click()
    Unload Me
End Sub

Private Sub Form_Load()
        'Create the connection

    Set conAVB = New ADODB.Connection
    
    conAVB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Persist Security Info=False;Data Source=" & App.Path & _
        "\AVB8.mdb;Mode = readwrite"
    conAVB.Open
End Sub
