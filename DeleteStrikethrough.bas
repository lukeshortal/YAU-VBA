Attribute VB_Name = "DelStrikethroughText"

Function HasStrike(Rng As Range) As Boolean
   HasStrike = Rng.Font.Strikethrough
End Function
