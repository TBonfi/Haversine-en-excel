Attribute VB_Name = "Module1"
Public Function distancia(latitude1, longitude1, latitude2, longitude2 As Double, Optional radius As Long = 6371)
Dim a, b, c As Variant
'6371km es el radio que us� siempre para puntos ubicados en Argentina, en el caso que est�s en otro pa�s o quieras m�s precisi�n pod�s buscar este dato en este lugar (no tengo relaci�n) https://rechneronline.de/earth-radius/
'6371km It's the radius I use for calculations all over Argentina, if you want to change this default or you mainly use lat/long from another country you can look up for this info in (Disclaimer, I'm affiliate to the site) https://rechneronline.de/earth-radius/

a = Application.WorksheetFunction.Radians((90 - latitude1))
b = Application.WorksheetFunction.Radians((90 - latitude2))
c = Application.WorksheetFunction.Radians((longitude1 - longitude2))

d = (WorksheetFunction.Acos(Cos(a) * Cos(b) + Sin(a) * Sin(b) * Cos(c))) * radius

distancia = d

End Function
