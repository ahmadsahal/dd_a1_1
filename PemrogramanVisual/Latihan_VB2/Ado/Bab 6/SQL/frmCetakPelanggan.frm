VERSION 5.00
Begin VB.Form frmCetakPelanggan 
   Caption         =   "Cetak Pelanggan"
   ClientHeight    =   840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   LinkTopic       =   "Form3"
   ScaleHeight     =   840
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBatal 
      Caption         =   "Batal"
      Height          =   492
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   1644
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "Cetak"
      Height          =   492
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1644
   End
End
Attribute VB_Name = "frmCetakPelanggan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim conAVB As ADODB.Connection
Dim rsPelanggan As ADODB.Recordset
Dim P As Printer

Private Sub cmdCetak_Click()
    CetakPelanggan
    Printer.EndDoc
End Sub
Private Sub CetakPelanggan()
    'Create the recordset
    Dim strsql As String
    Set rsPelanggan = New ADODB.Recordset
    strsql = "Select * from Pelanggan "
    rsPelanggan.Open strsql, conAVB, adOpenDynamic, adLockOptimistic, adCmdText
    'mencetak urut kode Pelanggan
    Dim MNo As Integer
    Dim Mhal As String
    Dim Mbaris As String
    Dim MGrs As String
    On Error GoTo SalahCetak
    With rsPelanggan
     'bawa pointer ke awal record yang dicetak
     .MoveFirst
     'tentukan font
     Printer.Font = "courier Times"
     
     'bawa head Printer ke awal halaman
     Printer.CurrentX = 0
     Printer.CurrentY = 0
     'mulai pengulangan
     MNo = 0
     Mhal = 0
     Do While Not rsPelanggan.EOF
        'cetak judul tabel
        Mhal = Mhal + 1
        Printer.Print "DAFTAR Pelanggan";
        Printer.Print Tab(60); "Hal :"; Format(Mhal, "###")
        MGrs = String$(70, "-")
        Printer.Print MGrs
        Printer.Print Tab(5); "No.";
        Printer.Print Tab(14); "Kode";
        Printer.Print Tab(25); "Nama";
        Printer.Print Tab(40); "Alamat";
        Printer.Print Tab(60); "Nomor"
        Printer.Print Tab(5);
        Printer.Print Tab(14); "Pelanggan";
        Printer.Print Tab(25); "Pelanggan";
        Printer.Print Tab(40); "Pelanggan";
        Printer.Print Tab(60); "Telepon"
        
        Printer.Print MGrs
        Mbaris = 0
        'mulai pengulangan cetak isi tabel
        Do While Not rsPelanggan.EOF And Mbaris <= 55
        MNo = MNo + 1
           
            Printer.Print Tab(5); MNo;
            Printer.Print Tab(14); ![Kode Pelanggan];
            Printer.Print Tab(25); ![Nama Pelanggan];
            Printer.Print Tab(40); ![ALamat Pelanggan];
            Printer.Print Tab(60); ![Telepon Pelanggan]
        Mbaris = Mbaris + 1
           .MoveNext
           If .EOF Then
              Exit Do
           End If
        Loop
       'Printer.Print MGrs
       Printer.NewPage
       If .EOF Then
          Exit Do
       End If
     Loop
     Printer.EndDoc

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
