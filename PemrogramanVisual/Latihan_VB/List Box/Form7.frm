VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7305
   LinkTopic       =   "Form7"
   ScaleHeight     =   5805
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   2280
      TabIndex        =   1
      Top             =   720
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Buka"
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   5040
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   6480
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Returns False if the command has been canceled, True otherwise.
Function SelectMultipleFiles(CD As CommonDialog, Filter As String, _
    Filenames() As String) As Boolean
    On Error GoTo ExitNow
    
    CD.Filter = "All files (*.*)|*.*|" & Filter
    CD.FilterIndex = 1
    CD.Flags = cdlOFNAllowMultiselect Or cdlOFNFileMustExist Or _
        cdlOFNExplorer
    CD.DialogTitle = "Select one or more files"
    CD.MaxFileSize = 10240
    CD.FileName = ""
    ' Exit if user presses Cancel.
    CD.CancelError = True
    CD.ShowOpen

    ' Parse the result to get filenames.
    Filenames() = Split(CD.FileName, vbNullChar)
    ' Signal success.
    SelectMultipleFiles = True
ExitNow:
End Function

Private Sub Command1_Click()
    Dim Filenames() As String, i As Integer
If SelectMultipleFiles(CD, "Text Dokumen (*.txt)|*.txt", Filenames()) Then
    If UBound(Filenames) = 0 Then
        ' The Filename property contained only one element.
     '   Print "Selected file: " & Filenames(0)
    Else
        ' The Filename property contained multiple elements.
       ' Print "Directory name: " & Filenames(0)
        For i = 1 To UBound(Filenames)
            Me.List1.AddItem Filenames(i)
        Next
    End If
End If
End Sub



 

