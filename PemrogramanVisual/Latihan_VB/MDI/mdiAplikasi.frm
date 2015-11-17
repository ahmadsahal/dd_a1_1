VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm mdiAplikasi 
   BackColor       =   &H8000000C&
   Caption         =   "Aplikasi"
   ClientHeight    =   4260
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7440
   Icon            =   "mdiAplikasi.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imlAplikasi 
      Left            =   2520
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAplikasi.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAplikasi.frx":0896
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAplikasi.frx":0CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAplikasi.frx":115E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAplikasi.frx":156E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   7440
      _CBHeight       =   390
      _Version        =   "6.7.8988"
      Child1          =   "tlbAplikasi"
      MinHeight1      =   330
      Width1          =   3450
      NewRow1         =   0   'False
      MinHeight2      =   330
      Width2          =   4260
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlbAplikasi 
         Height          =   330
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlAplikasi"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Input"
               Object.ToolTipText     =   "Input Data"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Output"
               Object.ToolTipText     =   "Output Data"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "spr"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Vertical"
               Object.ToolTipText     =   "Vertically"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Horizontal"
               Object.ToolTipText     =   "Horizontally"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cascade"
               Object.ToolTipText     =   "Cascade"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbAplikasi 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   3870
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   6085
            Text            =   "Created By Viansastra"
            TextSave        =   "Created By Viansastra"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   3228
            TextSave        =   "1/31/04"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   3228
            TextSave        =   "12:42 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuInput 
         Caption         =   "&Input Data"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuOutput 
         Caption         =   "&Output Data"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuspr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuHorizontal 
         Caption         =   "&Horizontal"
      End
      Begin VB.Menu mnuVertical 
         Caption         =   "&Vertical"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "mdiAplikasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    'mengambil nilai dari registry
    Me.Left = GetSetting("Aplikasi", "Definisi", "Left", 1000)
    Me.Top = GetSetting("Aplikasi", "Definisi", "Top", 1000)
    Me.Width = GetSetting("Aplikasi", "Definisi", "Width", 7600)
    Me.Height = GetSetting("Aplikasi", "Definisi", "Height", 6000)
    
    'menampilkan form logo
    frmLogo.Show
    
    'menampilkan pesan tertulis
    frmPesan.Pesan = "Aplikasi siap..."
    
    'menentukan waktu tampil dalam detik
    frmPesan.Lama = 3
End Sub

Private Sub MDIForm_Resize() 'jika MDI diubah ukuran
    'menentukan lebar dan tinggi frmlogo
    'agar menyesuaikan ukuran MDI
    With frmLogo
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'menyimpan nilai left, top,height dan width MDI pada registry
    SaveSetting "Aplikasi", "Definisi", "Left", Me.Left
    SaveSetting "Aplikasi", "Definisi", "Top", Me.Top
    SaveSetting "Aplikasi", "Definisi", "Height", Me.Height
    SaveSetting "Aplikasi", "Definisi", "Width", Me.Width
End Sub

Private Sub mnuClose_Click()
    'menutup form aktif
    On Error Resume Next
    Unload Me.ActiveForm
End Sub

Sub mnuInput_Click()
    'menampilkan pesan
    frmPesan.Pesan = "Menampilkan Input data"
    frmPesan.Lama = 3
    
    'menampilkan frmAnak1
    frmAnak1.Show
End Sub

Sub mnuOutput_Click()
    'menampilkan pesan
    frmPesan.Pesan = "Menampilkan Output data"
    frmPesan.Lama = 3
    
    'menampilkan frmAnak
    frmAnak2.Show
End Sub

Sub mnuCascade_Click()
    'menampilkan pesan
    frmPesan.Pesan = "Windows Cascade"
    frmPesan.Lama = 3
    
    'pengaturan tampilan MDIChild
    Me.Arrange vbCascade
End Sub

Sub mnuHorizontal_Click()
    'menampilkan pesan
    frmPesan.Pesan = "Windows Horizontal"
    frmPesan.Lama = 3
    
    'pengaturan tampilan MDIChild
    Me.Arrange vbHorizontal
End Sub

Sub mnuVertical_Click()
    'menampilkan pesan
    frmPesan.Pesan = "Windows Vertical"
    frmPesan.Lama = 3
    
    'pengaturan tampilan MDIChild
    Me.Arrange vbVertical
End Sub

Private Sub tlbAplikasi_ButtonClick(ByVal _
Button As MSComctlLib.Button)
    'menentukan tombol yang ditekan
    Select Case Button.Key
        Case "Input": Me.mnuInput_Click
        Case "Output": Me.mnuOutput_Click
        Case "Horizontal"
            Me.mnuHorizontal_Click
        Case "Vertical"
            Me.mnuVertical_Click
        Case "Cascade": Me.mnuCascade_Click
    End Select
End Sub
