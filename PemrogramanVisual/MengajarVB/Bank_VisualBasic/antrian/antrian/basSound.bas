Attribute VB_Name = "basSound"
Public Const SND_NOSTOP = &H10

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

  
  

   
