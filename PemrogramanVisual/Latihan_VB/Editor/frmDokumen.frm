VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDokumen 
   Caption         =   "frmDocument"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3075
   Icon            =   "frmDokumen.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2085
   ScaleWidth      =   3075
   Begin RichTextLib.RichTextBox rtfEditor 
      Height          =   1995
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3519
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmDokumen.frx":0442
   End
End
Attribute VB_Name = "frmDokumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub rtfEditor_SelChange()
    fFormUtama.tlbEditor.Buttons( _
    "Bold").Value = IIf(rtfEditor.SelBold, _
    tbrPressed, tbrUnpressed)
    fFormUtama.tlbEditor.Buttons( _
    "Italic").Value = IIf( _
    rtfEditor.SelItalic, _
    tbrPressed, tbrUnpressed)
    fFormUtama.tlbEditor.Buttons( _
    "Underline").Value = IIf( _
    rtfEditor.SelUnderline, _
    tbrPressed, tbrUnpressed)
    fFormUtama.tlbEditor.Buttons( _
    "Align Left").Value = IIf( _
    rtfEditor.SelAlignment = _
    rtfLeft, tbrPressed, tbrUnpressed)
    fFormUtama.tlbEditor.Buttons( _
    "Center").Value = IIf( _
    rtfEditor.SelAlignment = rtfCenter, _
    tbrPressed, tbrUnpressed)
    fFormUtama.tlbEditor.Buttons( _
    "Align Right").Value = IIf( _
    rtfEditor.SelAlignment = _
    rtfRight, tbrPressed, tbrUnpressed)
End Sub

Private Sub Form_Load()
    Form_Resize
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    rtfEditor.Move 100, 100, _
    Me.ScaleWidth - 200, Me.ScaleHeight - 200
    rtfEditor.RightMargin = _
    rtfEditor.Width - 400
End Sub

