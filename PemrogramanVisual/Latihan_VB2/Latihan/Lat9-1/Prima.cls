VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Kali"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mAwal As Long
Dim mAkhir As Long
Dim mHasil As Single

Public Event Selesai()

Public Property Let Awal(ByVal nilai As Long)
mAwal = nilai 'mengubah nilai
End Property

Public Property Get Awal() As Long
Awal = mAwal 'membaca nilai
End Property

Public Property Let Akhir(ByVal nilai As Long)

mAkhir = nilai 'mengubah nilai
End Property

Public Property Get Akhir() As Long
Akhir = mAkhir 'membaca nilai
End Property

Public Property Let Hasil(nilai As Single)
mHasil = nilai
End Property

Public Property Get Hasil() As Single
Hasil = mHasil
End Property

Public Sub Kali()
Hasil = mAwal * mAkhir
RaiseEvent Selesai
End Sub

Public Sub Pangkat(x As Integer, y As Integer)
Hasil = x ^ y
RaiseEvent Selesai
End Sub

Public Function Sisa(x As Integer, y As Integer) As Long
Sisa = x Mod y
Hasil = Sisa
RaiseEvent Selesai
End Function

Public Function Bagi() As Single
Bagi = Awal / Akhir
Hasil = Bagi
RaiseEvent Selesai
End Function
