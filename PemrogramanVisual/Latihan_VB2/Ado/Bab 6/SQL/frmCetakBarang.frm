VERSION 5.00
Begin VB.Form frmCetakBarang 
   Caption         =   "Cetak Barang"
   ClientHeight    =   1350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5370
   LinkTopic       =   "Form3"
   ScaleHeight     =   1350
   ScaleWidth      =   5370
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Arah Kertas"
      Height          =   492
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   4428
      Begin VB.OptionButton optLandscape 
         Caption         =   "Miring (Landscape)"
         Height          =   204
         Left            =   2400
         TabIndex        =   4
         Top             =   192
         Width           =   1932
      End
      Begin VB.OptionButton optPortrait 
         Caption         =   "Normal (Portrait)"
         Height          =   204
         Left            =   360
         TabIndex        =   3
         Top             =   192
         Width           =   2028
      End
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "Cetak"
      Height          =   492
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   1644
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "Batal"
      Height          =   492
      Left            =   2880
      TabIndex        =   0
      Top             =   720
      Width           =   1644
   End
End
Attribute VB_Name = "frmCetakBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim conAVB As ADODB.Connection
Dim rsBarang As ADODB.Recordset
Dim P As Printer

Private Sub cmdCetak_Click()
    Cetakbarang
    Printer.EndDoc
End Sub
Private Sub Cetakbarang()
    'Create the recordset
    Dim strsql As String
    Set rsBarang = New ADODB.Recordset
    strsql = "Select * from Barang "
    rsBarang.Open strsql, conAVB, adOpenDynamic, adLockOptimistic, adCmdText
    'mencetak urut kode barang
    Dim MNo As Integer
    Dim Mhal As String
    Dim Mbaris As String
    Dim MGrs As String
    On Error GoTo SalahCetak
    With rsBarang
     'bawa pointer ke awal record yang dicetak
     .MoveFirst
     'tentukan font
     Printer.Font = "Times New Roman"
     
     'bawa head Printer ke awal halaman
     Printer.CurrentX = 0
     Printer.CurrentY = 0
     'mulai pengulangan
     MNo = 0
     Mhal = 0
     Do While Not rsBarang.EOF
        'cetak judul tabel
        Mhal = Mhal + 1
        Printer.Print "DAFTAR BARANG";
        Printer.Print Tab(60); "Hal :"; Format(Mhal, "###")
        MGrs = String$(70, "-")
        Printer.Print MGrs
        Printer.Print Tab(5); "No.";
        Printer.Print Tab(10); "Kode";
        Printer.Print Tab(23); "Nama";
        Printer.Print Tab(40); "Jumlah";
        Printer.Print Tab(50); "Harga"
        Printer.Print MGrs
        Mbaris = 0
        'mulai pengulangan cetak isi tabel
        Do While Not rsBarang.EOF And Mbaris <= 55
        MNo = MNo + 1
           
           Printer.Print Tab(1); MNo;
           Printer.Print Tab(10); ![Kode Barang];
           Printer.Print Tab(23); ![Nama barang];
           Printer.Print Tab(40); ![Jumlah Barang];
           Printer.Print Tab(50); ![Harga Barang];
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
Private Sub Form_Activate()
    optPortrait.Value = True
    Printer.Orientation = vbPRORPortrait
End Sub

Private Sub optLandscape_Click()
    Printer.Orientation = vbPRORLandscape
End Sub

Private Sub optPortrait_Click()
    Printer.Orientation = vbPRORPortrait
End Sub
