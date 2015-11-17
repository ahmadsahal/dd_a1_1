Attribute VB_Name = "Module1"
Global MyDB As New ADODB.Connection
Global Cmd As New ADODB.Command
Global RcProfile As New ADOR.Recordset
Global RcSaving As New ADOR.Recordset
Global RcTransaksi As New ADOR.Recordset

Global Jumlah As Single

Global Waktu As Date
Global Stat As Boolean
Global User As String
Global Pass As String
Global JmlNasabah As Long
Global JudulTarik As String



Public Function GenVoucher() As String
Dim Temp(1 To 10) As Integer
Dim x As Integer

'generate angka acak 100 s/d 1000
For x = 1 To 10
    While Temp(x) < 100
        Temp(x) = Int(Rnd * 1000)
    Wend
    
    GenVoucher = GenVoucher + Str(Temp(x))
Next x

End Function

Public Function BukaDB(userx As String, passx As String) As Boolean
    
    'buka hubungan ke BANKRIA
    MyDB.Provider = "SQLOLEDB"
    MyDB.CursorLocation = adUseServer
    MyDB.Open "Server=IWAN;Database=BankRia;UID=" + userx + ";PWD=" + passx
    
    BukaDB = True
    
End Function
