VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCetak 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Grid"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCetak 
      Caption         =   "Cetak &Data"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   3960
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc adoData 
      Height          =   375
      Left            =   120
      Top             =   3480
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adoData"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid dgdData 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCetak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim myXL As Excel.Application
Dim myBK As Excel.Workbook
Dim myST As Excel.Worksheet
Dim Baris As Byte
Dim MSG As String

Private Sub cmdCetak_Click()
With adoData.Recordset
'jika data tidak kosong
If .RecordCount <> 0 Then
    MSG = MsgBox("Apakah printer Anda sudah siap", _
    32 + vbYesNo)
        If MSG = vbYes Then
            'Memanggil Aplikasi Excel .....
            Set myXL = New Excel.Application
            myXL.Visible = True
            'Membuka file:..\..\Cetak.xls"
            Set myBK = GetObject( _
            App.Path & "\Cetak.xls")
            'melakukan pengulangan
            'sampai batas EOF
            Do While Not .EOF
                'menentukan baris awal
                Baris = 4
                'menampilkan jendela Cetak.xls
                myBK.Windows("Cetak.xls").Visible = True
                'Mengaktifkan & memilih sheet "Sheet1"
                Set myST = myBK.Worksheets("Sheet1")
                myST.Select
                'mengaktifkan range
                myST.Range("A4:C33").Select
                'mengkosongkan range
                myST.Range("A4:C33").ClearContents
                'melakukan pengulangan
                'sampai batas EOF
                'dan Baris tidak melebihi 33
                Do While Not .EOF And Baris <= 33
                'mengisi kolom A dengan string
                'posisi recordset
                myST.Range("A" + CStr(Baris)).Value = _
                CStr(.AbsolutePosition)
                'mengisi kolom B dengan string
                'data pada field Nama
                myST.Range("B" + CStr(Baris)).Value = ![Nama]
                'mengisi kolom C dengan string
                'data pada field Alamat
                myST.Range("C" + CStr(Baris)).Value = ![Alamat]
                'melanjutkan pengisian baris
                Baris = Baris + 1
                'maju ke posisi data berikutnya
                .MoveNext
            Loop
                'mencetak ke kertas
                myST.PrintOut
                'jika baris melebihi 33
                If Baris > 33 Then
                    MSG = MsgBox("SIAPKAN KERTAS BARU", 4 + 64)
                    If MSG = vbNo Then
                        MsgBox "Pencetakan Selesai", vbInformation
                        Exit Do
                    End If
                End If
        Loop
            'jika mencapai EOF
                If .EOF Then
                    MsgBox "Pencetakan selesai", vbInformation
                'menutup jendela Cetak.xls
                myBK.Close False
                'keluar dari aplikasi Excel
                myXL.Quit
                'pengosongan
                Set myXL = Nothing
                Set myBK = Nothing
                Set myST = Nothing
            'jika mencapai EOF
            End If
        End If
    'jika data kosong
    Else: MsgBox "Data masih kosong", vbExclamation
End If
End With
End Sub

Private Sub Form_Load()
    With Me.adoData
        .ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & _
        App.Path & "\data.mdb;Mode=ReadWrite;" & _
        "Persist Security Info=False"
        .CommandType = adCmdText
        .RecordSource = _
        "SELECT * FROM tblTemanku order by [Nama]"
        'order by digunakan untuk mengurutkan data
        'order by [Nama]")untuk mengurutkan data dengan
        'patokan pada field Nama
        .Refresh
  End With
      Set Me.dgdData.DataSource = Me.adoData
End Sub



