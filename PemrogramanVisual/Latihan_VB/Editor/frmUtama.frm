VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm frmUtama 
   BackColor       =   &H8000000C&
   Caption         =   "Editor File String"
   ClientHeight    =   4155
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7725
   Icon            =   "frmUtama.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComCtl3.CoolBar clbEditor 
      Align           =   1  'Align Top
      Height          =   750
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   1323
      BandCount       =   2
      _CBWidth        =   7725
      _CBHeight       =   750
      _Version        =   "6.0.8169"
      Child1          =   "tlbEditor"
      MinHeight1      =   330
      Width1          =   1215
      NewRow1         =   0   'False
      Child2          =   "tlbFont"
      MinHeight2      =   330
      Width2          =   1440
      NewRow2         =   -1  'True
      Begin MSComctlLib.Toolbar tlbFont 
         Height          =   330
         Left            =   165
         TabIndex        =   3
         Top             =   390
         Width           =   7470
         _ExtentX        =   13176
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlEditor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageKey        =   "Font"
               Style           =   5
               Object.Width           =   1e-4
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
         Begin VB.ComboBox cboUkuran 
            Height          =   315
            Left            =   720
            TabIndex        =   4
            Top             =   0
            Width           =   615
         End
      End
      Begin MSComctlLib.Toolbar tlbEditor 
         Height          =   330
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   7470
         _ExtentX        =   13176
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlEditor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "New"
               Object.ToolTipText     =   "1024"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Open"
               Object.ToolTipText     =   "1025"
               ImageKey        =   "Open"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               Object.ToolTipText     =   "1026"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Copy"
               Object.ToolTipText     =   "1027"
               ImageKey        =   "Copy"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Paste"
               Object.ToolTipText     =   "1028"
               ImageKey        =   "Paste"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cut"
               Object.ToolTipText     =   "1029"
               ImageKey        =   "Cut"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bold"
               Object.ToolTipText     =   "1030"
               ImageKey        =   "Bold"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Italic"
               Object.ToolTipText     =   "1031"
               ImageKey        =   "Italic"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Underline"
               Object.ToolTipText     =   "1032"
               ImageKey        =   "Underline"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Align Left"
               Object.ToolTipText     =   "1033"
               ImageKey        =   "Align Left"
               Style           =   2
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Center"
               Object.ToolTipText     =   "1034"
               ImageKey        =   "Center"
               Style           =   2
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Align Right"
               Object.ToolTipText     =   "1035"
               ImageKey        =   "Align Right"
               Style           =   2
            EndProperty
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog dlgEditor 
      Left            =   2760
      Top             =   1350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      InitDir         =   "C:\Master\Editor"
   End
   Begin MSComctlLib.StatusBar stbEditor 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   3885
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   3069
            MinWidth        =   3069
            Text            =   "Created By Viansastra"
            TextSave        =   "Created By Viansastra"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "19/07/04"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "10:35"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlEditor 
      Left            =   1740
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":0442
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":0554
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":0666
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":0778
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":088A
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":099C
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":0AAE
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":0BC0
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":0CD2
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":0DE4
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":0EF6
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":1008
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":111A
            Key             =   "Fon"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":145A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":156E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":1682
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":1796
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":18AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":1A46
            Key             =   "Font"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "1000"
      Begin VB.Menu mnuFileNew 
         Caption         =   "1001"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "1002"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "1003"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "1004"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "1005"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "1006"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "1007"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "1008"
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "1009"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "1010"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "1011"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "1012"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "1013"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "1014"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "1016"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "1017"
      End
      Begin VB.Menu mnuWindowBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "1018"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "1019"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "1020"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "1021"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "1022"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "1023"
      End
   End
End
Attribute VB_Name = "frmUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboUkuran_Click()
On Error Resume Next
    Me.ActiveForm.rtfEditor.SelFontSize = _
    Val(Me.cboUkuran.Text)
End Sub

Private Sub cboUkuran_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        On Error Resume Next
    Me.ActiveForm.rtfEditor.SelFontSize = _
    Val(Me.cboUkuran.Text)
    End If
End Sub

Private Sub MDIForm_Load()
    AmbilNilaiRes Me
    Me.Left = GetSetting(App.Title, _
    "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, _
    "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, _
    "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, _
    "Settings", "MainHeight", 6500)
    LoadDokumenBaru
    LoadHurufDanUkuran
End Sub

Sub LoadHurufDanUkuran(): 'On Error Resume Next
    Dim j As Long, k As Long
    For j = 1 To Screen.FontCount
        Me.tlbFont.Buttons(1).ButtonMenus.Add j, _
        "Huruf" & CStr(j), Screen.Fonts(j)
    Next
    For k = 1 To 72
        Me.cboUkuran.AddItem CStr(k)
    Next
End Sub

Private Sub LoadDokumenBaru()
    Static lDokumenCount As Long
    Dim frmD As frmDokumen
    lDokumenCount = lDokumenCount + 1
    Set frmD = New frmDokumen
    frmD.Caption = "Dokumen " & lDokumenCount
    frmD.Show
End Sub


Private Sub MDIForm_Resize()
On Error Resume Next
    With frmLogo
        .Move 0, 0, ScaleWidth, _
        ScaleHeight
        .imgLogo.Move (.ScaleWidth - _
        .imgLogo.Width) / 2, (.ScaleHeight - _
        .imgLogo.Height) / 2
    End With
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", _
        "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", _
        "MainTop", Me.Top
        SaveSetting App.Title, "Settings", _
        "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", _
        "MainHeight", Me.Height
    End If
End Sub

Private Sub tlbEditor_ButtonClick(ByVal _
Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            LoadDokumenBaru
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Bold"
            ActiveForm.rtfEditor.SelBold = _
            Not ActiveForm.rtfEditor.SelBold
            Button.Value = _
            IIf(ActiveForm.rtfEditor.SelBold, _
            tbrPressed, tbrUnpressed)
        Case "Italic"
            ActiveForm.rtfEditor.SelItalic = _
            Not ActiveForm.rtfEditor.SelItalic
            Button.Value = _
            IIf(ActiveForm.rtfEditor.SelItalic, _
            tbrPressed, tbrUnpressed)
        Case "Underline"
            ActiveForm.rtfEditor.SelUnderline = _
            Not ActiveForm.rtfEditor.SelUnderline
            Button.Value = _
            IIf(ActiveForm.rtfEditor.SelUnderline, _
            tbrPressed, tbrUnpressed)
        Case "Align Left"
            ActiveForm.rtfEditor.SelAlignment = _
            rtfLeft
        Case "Center"
            ActiveForm.rtfEditor.SelAlignment = _
            rtfCenter
        Case "Align Right"
            ActiveForm.rtfEditor.SelAlignment = _
            rtfRight
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowNewWindow_Click()
    LoadDokumenBaru
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not _
    mnuViewStatusBar.Checked
    stbEditor.Visible = _
    mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not _
    mnuViewToolbar.Checked
    tlbEditor.Visible = _
    mnuViewToolbar.Checked
End Sub

Private Sub mnuEditPaste_Click()
    On Error Resume Next
    ActiveForm.rtfEditor.SelRTF = _
    Clipboard.GetText

End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfEditor.SelRTF

End Sub

Private Sub mnuEditCut_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfEditor.SelRTF
    ActiveForm.rtfEditor.SelText = vbNullString

End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileSaveAs_Click()
    Dim sFile As String
    If ActiveForm Is Nothing Then Exit Sub
    With dlgEditor
        .DialogTitle = "Save As"
        .CancelError = False
        .Filter = _
        "Teks Dokumen (*.txt)|*.txt|" & _
        "Semua Files (*.*)|*.*|"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    ActiveForm.Caption = sFile
    ActiveForm.rtfEditor.SaveFile sFile
End Sub

Private Sub mnuFileSave_Click()
    Dim sFile As String
    If Left$(ActiveForm.Caption, 9) = _
    "Dokumen 1" Then
        With dlgEditor
            .DialogTitle = "Simpan"
            .CancelError = False
            .Filter = _
            "Teks Dokumen (*.txt)|*.txt|" & _
            "Semua Files (*.*)|*.*|"
            .ShowSave
            If Len(.FileName) = 0 Then
                Exit Sub
            End If
            sFile = .FileName
        End With
        ActiveForm.rtfEditor.SaveFile sFile
    Else
        sFile = ActiveForm.Caption
        ActiveForm.rtfEditor.SaveFile sFile
    End If
End Sub

Private Sub mnuFileClose_Click()
On Error Resume Next
    Unload ActiveForm
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String
    If ActiveForm Is Nothing Then LoadDokumenBaru
    With dlgEditor
        .DialogTitle = "Buka"
        .CancelError = False
        .Filter = _
        "Teks Dokumen (*.txt)|*.txt|" & _
        "Semua Files (*.*)|*.*|"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    ActiveForm.rtfEditor.LoadFile sFile
    ActiveForm.Caption = sFile
End Sub

Private Sub mnuFileNew_Click()
    LoadDokumenBaru
End Sub

Private Sub tlbFont_ButtonMenuClick(ByVal _
ButtonMenu As MSComctlLib.ButtonMenu):
On Error Resume Next
    Me.ActiveForm.rtfEditor.SelFontName = _
    ButtonMenu.Text
End Sub
