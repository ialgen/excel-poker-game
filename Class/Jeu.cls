VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "jeu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Base 1

Private joueurs As Collection
Private joueurs_en_jeu As Collection
Private paquet As Collection
Public pot As Long
Public blind As Integer

Property Let let_joueurs(nb_joueurs As Integer, argent_ As Long)
    Set joueurs_ = New Collection
    For i = 1 To CInt(nb_joueurs)
        Set joueur_ = New joueur
        Let joueur_.let_stack = argent_
        joueurs_.Add (joueur_)
    Next i
    Set joueurs = joueurs_
End Property
Property Get get_joueurs()
    Set get_joueurs = joueurs
End Property
Property Let let_joueurs_en_jeu(joueurs_en_jeu_ As Collection)
    Set joueurs_en_jeu = joueurs_en_jeu_
End Property
Property Get get_joueurs_en_jeu()
    Set get_joueurs_en_jeu = joueurs_en_jeu
End Property

Property Let let_paquet(paquet_ As Collection)
    Set paquet = paquet_
End Property
Property Get get_paquet()
    Set get_paquet = paquet
End Property

Private Sub Class_Initialize()
    'Création d'une collection de cartes ordonnees, pas de doublons possibles
    Set cartes_triees_ = New Collection
    For i = 1 To 13
        For j = 1 To 4
            Set carte = New carte
            carte.let_valeur = i
            carte.let_couleur = j
            cartes_triees_.Add (carte)
        Next j
    Next i

    'Melange de la collection de cartes sans doublons
    Set cartes_melangees_ = New Collection
    For Each item In cartes_triees_
        cartes_melangees_.Add item
    Next
    For i = 1 To 52
        nb = WorksheetFunction.RandBetween(1, 52)
        
        Set carte_temp = New carte
        Let carte_temp.let_couleur = cartes_melangees_.item(nb).get_couleur
        Let carte_temp.let_valeur = cartes_melangees_.item(nb).get_valeur
        
        Let cartes_melangees_.item(nb).let_couleur = cartes_melangees_.item(i).get_couleur
        Let cartes_melangees_.item(nb).let_valeur = cartes_melangees_.item(i).get_valeur

        Let cartes_melangees_.item(i).let_couleur = carte_temp.get_couleur
        Let cartes_melangees_.item(i).let_valeur = carte_temp.get_valeur
    Next i

    Let Me.let_paquet = cartes_melangees_
End Sub

Sub melange()

    'Melange de la collection de cartes sans doublons
    Set cartes_melangees_ = New Collection
    For Each item In Me.get_paquet
        cartes_melangees_.Add item
    Next
    For i = 1 To 52
        nb = WorksheetFunction.RandBetween(1, 52)
        
        Set carte_temp = New carte
        Let carte_temp.let_couleur = cartes_melangees_.item(nb).get_couleur
        Let carte_temp.let_valeur = cartes_melangees_.item(nb).get_valeur
        
        Let cartes_melangees_.item(nb).let_couleur = cartes_melangees_.item(i).get_couleur
        Let cartes_melangees_.item(nb).let_valeur = cartes_melangees_.item(i).get_valeur

        Let cartes_melangees_.item(i).let_couleur = carte_temp.get_couleur
        Let cartes_melangees_.item(i).let_valeur = carte_temp.get_valeur
    Next i

    Let Me.let_paquet = cartes_melangees_

End Sub
