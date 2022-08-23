VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "carte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private couleur As Integer
Private valeur As Integer

'Property couleur en string
Property Get get_couleur_string()
    Select Case couleur
        Case 1
            get_couleur_string = "COEUR"
        Case 2
            get_couleur_string = "CARREAU"
        Case 3
            get_couleur_string = "TREFLE"
        Case 4
            get_couleur_string = "PIQUE"
        Case Else
    End Select
End Property

'Property valeur en string
Property Get get_valeur_string()
    Select Case valeur
        Case 13
            get_valeur_string = "A"
        Case 12
            get_valeur_string = "K"
        Case 11
            get_valeur_string = "Q"
        Case 10
            get_valeur_string = "J"
        Case 9
            get_valeur_string = 10
        Case 8
            get_valeur_string = 9
        Case 7
            get_valeur_string = 8
        Case 6
            get_valeur_string = 7
        Case 5
            get_valeur_string = 6
        Case 4
            get_valeur_string = 5
        Case 3
            get_valeur_string = 4
        Case 2
            get_valeur_string = 3
        Case 1
            get_valeur_string = 2
        Case Else
    End Select
End Property

'Property couleur en numerique
Property Let let_couleur(couleur_ As Integer)
    If couleur_ < 1 Or couleur_ > 4 Then
        Debug.Print "ERROR : i_couleur is not between 1 and 4"
    Else
        couleur = couleur_
    End If
End Property

Property Get get_couleur()
    get_couleur = couleur
End Property

'Property valeur en numerique
Property Let let_valeur(valeur_ As Integer)
    If valeur_ < 1 Or valeur_ > 13 Then
        Debug.Print "ERROR : valeur_ is not between 1 and 13"
    Else
        valeur = valeur_
    End If
End Property

Property Get get_valeur()
    get_valeur = valeur
End Property
