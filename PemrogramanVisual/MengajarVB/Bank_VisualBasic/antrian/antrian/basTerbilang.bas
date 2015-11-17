Attribute VB_Name = "basTerbilang"
Public Function Bilang(Value As Long) As String
   
   Select Case Value
      Case 0: Bilang = ""
      Case 1: Bilang = " Satu"
      Case 2: Bilang = " Dua"
      Case 3: Bilang = " Tiga"
      Case 4: Bilang = " Empat"
      Case 5: Bilang = " Lima"
      Case 6: Bilang = " Enam"
      Case 7: Bilang = " Tujuh"
      Case 8: Bilang = " Delapan"
      Case 9: Bilang = " Sembilan"
      Case 10: Bilang = " Sepuluh"
      Case 11: Bilang = " Sebelas"
      Case 12 To 19: Bilang = Bilang(Value Mod 10) & " Belas"
      Case 20 To 99: Bilang = Bilang(Int(Value / 10)) & " Puluh" & Bilang(Value Mod 10)
      Case 100 To 199: Bilang = " Seratus" & Bilang(Value Mod 100)
      Case 200 To 999: Bilang = Bilang(Int(Value / 100)) & " Ratus" & Bilang(Value Mod 100)
   End Select

End Function
