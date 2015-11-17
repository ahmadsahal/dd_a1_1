VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmADO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adodc"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc adoTeman 
      Height          =   330
      Left            =   720
      Top             =   2160
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\HOUSE OF INOCHI\MASTER OF INOCHI\Master\Ado\data.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\HOUSE OF INOCHI\MASTER OF INOCHI\Master\Ado\data.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblTemanku"
      Caption         =   "adoTeman"
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
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdMundur 
      Caption         =   "M&undur"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdMaju 
      Caption         =   "&Maju"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdAkhir 
      Caption         =   "A&khir"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdAwal 
      Caption         =   "&Awal"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtAlamat 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtNama 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "A&lamat"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "&Nama"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmAdo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub TampilkanData()
    With adoTeman.Recordset
        Me.txtNama.Text = !Nama
        Me.txtAlamat.Text = !Alamat
    End With
End Sub

Private Sub cmdAwal_Click()
    With adoTeman.Recordset
        .MoveFirst
        If .RecordCount <> 0 Then
            Call TampilkanData
        Else: MsgBox "Data masih kosong"
        End If
    End With
End Sub

Private Sub cmdAkhir_Click()
    With adoTeman.Recordset
        .MoveLast
        If .RecordCount <> 0 Then
            Call TampilkanData
        Else: MsgBox "Data masih kosong"
        End If
    End With
End Sub

Private Sub cmdMaju_Click()
    With adoTeman.Recordset
        .MoveNext
        If .EOF Then .MoveLast
        If .RecordCount <> 0 Then
            Call TampilkanData
        Else: MsgBox "Data masih kosong"
        End If
    End With
End Sub

Private Sub cmdMundur_Click()
    With adoTeman.Recordset
        .MovePrevious
        If .BOF Then .MoveFirst
        If .RecordCount <> 0 Then
            Call TampilkanData
        Else: MsgBox "Data masih kosong"
        End If
    End With
End Sub

Private Sub cmdSimpan_Click()
    With adoTeman.Recordset
        .AddNew
        !Nama = Me.txtNama.Text
        !Alamat = Me.txtAlamat.Text
        .Update
    End With
End Sub

Private Sub cmdHapus_Click()
    With adoTeman.Recordset
        If .RecordCount <> 0 Then
            .Delete
            .MoveNext
            If .EOF Then .MovePrevious
            If .BOF Then MsgBox "Data" & _
            " sudah kosong"
        Else: MsgBox "Data masih kosong"
        End If
    End With
End Sub

Private Sub Form_Load()
    adoTeman.Refresh
End Sub


