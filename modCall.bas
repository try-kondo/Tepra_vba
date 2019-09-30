

' モジュール名の指定
Attribute VB_Name =  "modCall"


' Excel内のオブジェクトに登録する用のマクロ

Public sub CallBihin()
	
	Call modBihin.Bihin_main

End sub

Public sub CallAtesaki()
	
	Call modAtesaki.Atesaki_main
	
End sub