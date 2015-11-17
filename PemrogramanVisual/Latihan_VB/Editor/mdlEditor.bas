Attribute VB_Name = "mdlEditor"
Public fFormUtama As frmUtama

Sub Main()
    Set fFormUtama = New frmUtama
    Load fFormUtama
    Load frmLogo
    frmLogo.Show
    fFormUtama.Show
End Sub

Sub AmbilNilaiRes(frm As Form)
    On Error Resume Next
    Dim Kontrol As Control
    Dim Objek As Object
    Dim TipeKontrol As String
    Dim xNilai As Integer
    frm.Caption = LoadResString(CInt(frm.Tag))
    For Each Kontrol In frm.Controls
        Set Kontrol.Font = fnt
        TipeKontrol = TypeName(Kontrol)
        If TipeKontrol = "Label" Then
            Kontrol.Caption = LoadResString( _
            CInt(Kontrol.Tag))
        ElseIf TipeKontrol = "Menu" Then
            Kontrol.Caption = LoadResString( _
            CInt(Kontrol.Caption))
        ElseIf TipeKontrol = "TabStrip" Then
            For Each Objek In Kontrol.Tabs
                Objek.Caption = LoadResString( _
                CInt(Objek.Tag))
                Objek.ToolTipText = _
                LoadResString(CInt( _
                Objek.ToolTipText))
            Next
        ElseIf TipeKontrol = "Toolbar" Then
            For Each Objek In Kontrol.Buttons
                Objek.ToolTipText = LoadResString( _
                CInt(Objek.ToolTipText))
            Next
        ElseIf TipeKontrol = "ListView" Then
            For Each Objek In Kontrol _
            .ColumnHeaders
                Objek.Text = LoadResString( _
                CInt(Objek.Tag))
            Next
        Else
            xNilai = 0
            xNilai = Val(Kontrol.Tag)
            If xNilai > 0 Then Kontrol.Caption = _
            LoadResString(xNilai)
            xNilai = 0
            xNilai = Val(Kontrol.ToolTipText)
            If xNilai > 0 Then Kontrol _
            .ToolTipText = LoadResString(xNilai)
        End If
    Next
End Sub


