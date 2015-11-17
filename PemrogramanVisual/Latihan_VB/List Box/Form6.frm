VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3165
   LinkTopic       =   "Form6"
   ScaleHeight     =   2460
   ScaleWidth      =   3165
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2280
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form6.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form6.frx":0454
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageCombo imgCombo 
      Height          =   330
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      ImageList       =   "ImageList1"
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub form_load()
    Dim fso As New Scripting.FileSystemObject, dr As Scripting.Drive
    Dim drLabel As String, drImage As String
    ' Assume that the ImageCombo control is linked to an ImageList
    ' control that includes three icons with the following key names.
    imgCombo.ComboItems.Add , , "My Computer", 1
    For Each dr In fso.Drives
        ' Use a different image for each type of drive.
        Select Case dr.DriveType
            Case Removable:  drImage = "FloppyDrive"
            Case CDRom:      drImage = "CDDrive"
            Case Else:       drImage = "HardDrive"
        End Select
        ' Retrieve the letter and (if possible) the volume label.
        drLabel = dr.DriveLetter & ": "
        If dr.IsReady Then
            If Len(dr.VolumeName) Then drLabel = drLabel & "[" & _
                dr.VolumeName & "]"
        End If
        ' Add an indented item to the combo.
        imgCombo.ComboItems.Add , dr.DriveLetter, drLabel, 2, , 2
    Next
    ' Select the current drive.
    Set imgCombo.SelectedItem = imgCombo.ComboItems(Left$(CurDir$, 1))
End Sub

