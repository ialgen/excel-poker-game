VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "hand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private carte_1 As carte
Private carte_2 As carte

Property Let let_carte_1(carte_ As carte)
   Set carte_1 = carte_
End Property
Property Let let_carte_2(carte_ As carte)
   Set carte_2 = carte_
End Property
Property Get get_carte_1()
   Set get_carte_1 = carte_1
End Property
Property Get get_carte_2()
   Set get_carte_2 = carte_2
End Property
