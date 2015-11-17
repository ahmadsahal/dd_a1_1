Attribute VB_Name = "BasTerbilang"
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
      
      Case 1000 To 1999: Bilang = " Seribu" & Bilang(Value Mod 1000)
      Case 2000 To 9999: Bilang = Bilang(Int(Value / 1000)) & " Ribu" & Bilang(Value Mod 1000)
      
      Case 10000 To 19999: Bilang = " Sepuluh Ribu" & Bilang(Value Mod 10000)
      Case 20000 To 99999: Bilang = Bilang(Int(Value / 10000)) & " Puluh Ribu" & Bilang(Value Mod 10000)
      
      Case 100000 To 199999: Bilang = " Seratus Ribu" & Bilang(Value Mod 100000)
      Case 200000 To 999999: Bilang = Bilang(Int(Value / 100000)) & " Ratus Ribu" & Bilang(Value Mod 100000)
      
      Case 1000000 To 1999999: Bilang = " Satu Juta" & Bilang(Value Mod 1000000)
      Case 2000000 To 9999999: Bilang = Bilang(Int(Value / 1000000)) & " Juta" & Bilang(Value Mod 1000000)
      
   End Select

End Function

