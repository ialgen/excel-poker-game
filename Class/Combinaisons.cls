VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Combinaisons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Base 1

'Fonction qui renvoie le nom associé un numero de 1 a 10
'ce dernier sera obtenu grace a la fonction best_main

Function type_best_main(numero As Integer) As String

    Set hierarchie_main = New Collection

    hierarchie_main.Add "Quinte_Flush_Royale"
    hierarchie_main.Add "Quinte_Flush"
    hierarchie_main.Add "Carre"
    hierarchie_main.Add "Full"
    hierarchie_main.Add "Couleur"
    hierarchie_main.Add "Suite"
    hierarchie_main.Add "Brelan"
    hierarchie_main.Add "Double_Paire"
    hierarchie_main.Add "Paire"
    hierarchie_main.Add "Hauteur"
    
    type_best_main = hierarchie_main.item(numero)

End Function

'Fonction qui renvoie la meilleure main d'un jeu de cartes via un Array
'le premier element est le numero associe a une combinaison
'le second element renvoie les donnees associees a cette combinaisons (valeur de la paire ... )

Function best_main(jeu_cartes As Scripting.Dictionary)

    Dim hierarchie As Scripting.Dictionary
    Set hierarchie = New Scripting.Dictionary

    hierarchie.Add 1, quinte_flush_royale(jeu_cartes)
    hierarchie.Add 2, quinte_flush(jeu_cartes)
    hierarchie.Add 3, carre(jeu_cartes)
    hierarchie.Add 4, full(jeu_cartes)
    hierarchie.Add 5, couleur(jeu_cartes)
    hierarchie.Add 6, suite(jeu_cartes)
    hierarchie.Add 7, brelan(jeu_cartes)
    hierarchie.Add 8, double_paire(jeu_cartes)
    hierarchie.Add 9, paire(jeu_cartes)
    hierarchie.Add 10, hauteur(jeu_cartes)

    'plus la  combinaison du jeu de cartes est forte plus iter est faible (plus proche de 1 que de 10)
    For iter = 1 To 10
        If hierarchie(iter)(1) <> 0 Then
            best_main = Array(iter, hierarchie(iter), "")
            GoTo Last
        End If
    Next iter

Last:

End Function

'Fonction qui renvoie si un jeu de cartes a une quinte_flush_royale via un Array
'Si non renvoie 0
'Si oui renvoie la hauteur de la quinte flush_royale

Function quinte_flush_royale(jeu_cartes As Scripting.Dictionary)
    Dim hauteur_ As Integer
    Dim compteur As Integer
    Dim issuite As Boolean
    quinte_flush_royale = Array(0)
    
    For Each Key In dict_valeur_couleur(jeu_cartes).Keys
        
        hauteur_ = 0
        compteur = 0
        
        'il suffit de parcourir les 5 meilleures valeurs
        For iter = 13 To 9 Step -1
            If dict_valeur_couleur(jeu_cartes)(Key)(iter) >= 1 Then
                compteur = compteur + 1
                If compteur = 1 Then hauteur_ = iter
            Else
                compteur = 0
                hauteur_ = 0
            End If
            'il y a une quinte flush royale
            If compteur = 5 Then
                quinte_flush_royale = Array(hauteur_)
                GoTo Fin:
            End If
        Next iter
    Next Key
Fin:

End Function

'Fonction qui renvoie si un jeu de cartes a une quinte_flush via un Array
'Si non renvoie 0
'Si oui renvoie la hauteur de la quinte flush

Function quinte_flush(jeu_cartes As Scripting.Dictionary)
    
    Dim hauteur_ As Integer
    Dim compteur As Integer
    Dim issuite As Boolean
    quinte_flush = Array(0)
    
     For Each Key In dict_valeur_couleur(jeu_cartes).Keys
        
        hauteur_ = 0
        compteur = 0
    
        For iter = 13 To 1 Step -1
            If dict_valeur_couleur(jeu_cartes)(Key)(iter) >= 1 Then
                compteur = compteur + 1
                If compteur = 1 Then hauteur_ = iter
            Else
                compteur = 0
                hauteur_ = 0
            End If
        
            If compteur = 5 Then
                quinte_flush = Array(hauteur_)
                GoTo Fin:
            End If
        Next iter
    Next Key
Fin:
End Function

'Fonction qui renvoie si un jeu de cartes a un carre via un Array
'Si non renvoie 0
'Si oui renvoie la valeur du carre et la hauteur

Function carre(jeu_cartes As Scripting.Dictionary)
    Dim hauteur_ As Integer
    Dim cle_carre As Integer
    Dim iscarre As Boolean
    
    iscarre = False
    For Each Key In dict_valeur(jeu_cartes).Keys
    
        If (dict_valeur(jeu_cartes)(Key) < 4 And dict_valeur(jeu_cartes)(Key) >= 1) Then
            hauteur_ = Key
        End If
        
        If dict_valeur(jeu_cartes)(Key) = 4 Then
            iscarre = True
            cle_carre = Key
        End If
        
    Next Key

    If iscarre = True Then
        carre = Array(cle_carre, hauteur_)
    Else
        carre = Array(0)
    End If
    
End Function

'Fonction qui renvoie si un jeu de cartes a un full via un Array
'Si non renvoie 0
'Si oui renvoie la valeur du brelan et de la paire

Function full(jeu_cartes As Scripting.Dictionary)
    
    Dim isbrelan As Boolean
    Dim ispaire As Boolean
    Dim cle_brelan As Integer
    Dim cle_paire As Integer
    Dim Col_valeur As Scripting.Dictionary
    Set Col_valeur = dict_valeur(jeu_cartes)
    
    cle_brelan = 0
    cle_paire = 0
    ispaire = False
    isbrelan = False
    
    'le meilleur brelan
    For Each Key In Col_valeur.Keys
    
        If Col_valeur(Key) = 3 And Key > cle_brelan Then
            cle_brelan = Key
            isbrelan = True
        End If
        
    Next Key
    
    If cle_brelan = 0 Then GoTo Fin:
    
    'enlever la cle brelan
    Col_valeur.Remove (cle_brelan)
    
    'la meilleure paire
    For Each Key In Col_valeur.Keys
    
        If Col_valeur(Key) >= 2 And Key > cle_paire Then
            cle_paire = Key
            ispaire = True
        End If
        
    Next Key
    
    
Fin:
    If ispaire = True And isbrelan = True Then
        full = Array(cle_brelan, cle_paire)
    Else
        full = Array(0)
    End If
    
End Function

'Fonction qui renvoie si un jeu de cartes a une couleur via un Array
'Si non renvoie 0
'Si oui renvoie la valeur des cartes formant la couleur

Function couleur(jeu_cartes As Scripting.Dictionary)
    Dim cle_couleur As String
    Dim iscouleur As Boolean
    Dim cle_dict As Scripting.Dictionary
    Set cle_dict = New Scripting.Dictionary
    iscouleur = False
    
    'Determination de la couleur (COEUR ou CARREAU etc)
    For Each Key In dict_couleur(jeu_cartes).Keys
        
        If dict_couleur(jeu_cartes)(Key) >= 5 Then
            iscouleur = True
            cle_couleur = Key
        End If
        
    Next Key
    
    'Si pas de couleur aller a la fin
    If iscouleur = False Then
        GoTo Fin2
    End If
    
    'Determination du dictionnaire regroupant les cartes de la couleur
    For Each Key In dict_valeur_couleur(jeu_cartes).Keys
        If Key = cle_couleur Then
            Set cle_dict = dict_valeur_couleur(jeu_cartes)(Key)
        End If
    Next Key
    
    Dim couleur_val_vect As Variant
    ReDim couleur_val_vect(1, 5) As Variant
    hauteur_test = 0
    
    'Determination de l'ordre decroissant des cartes de notre couleur
    For iter = 1 To 5
    
        For Each Key In cle_dict.Keys
    
            If (Key > hauteur_test And cle_dict(Key) = 1) Then
                hauteur_test = Key
            End If
        
        Next Key
        
        'enlever la hauteur
        cle_dict.Remove (hauteur_test)
        couleur_val_vect(1, iter) = hauteur_test
        hauteur_test = 0
        
    Next iter
    
Fin2:
    If iscouleur = True Then
        couleur = Array(couleur_val_vect(1, 1), couleur_val_vect(1, 2), couleur_val_vect(1, 3), couleur_val_vect(1, 4), couleur_val_vect(1, 5))
    Else
        couleur = Array(0)
    End If
    
End Function

'Fonction qui renvoie si un jeu de cartes a une suite via un Array
'Si non renvoie 0
'Si oui renvoie les valeurs des cartes de la suite

Function suite(jeu_cartes As Scripting.Dictionary)
    Dim hauteur_ As Integer
    Dim compteur As Integer
    Dim issuite As Boolean
    
    hauteur_ = 0
    compteur = 0
    issuite = False
    
    For iter = 13 To 1 Step -1
        If dict_valeur(jeu_cartes)(iter) >= 1 Then
            compteur = compteur + 1
            If compteur = 1 Then hauteur_ = iter
        Else
            compteur = 0
            hauteur_ = 0
        End If
        
        'il y a une suite
        If compteur = 5 Then
            issuite = True
            GoTo Fin:
        End If
    Next iter
    
Fin:
    If issuite = True Then
        suite = Array(hauteur_)
    Else
        suite = Array(0)
    End If
    
End Function

'Fonction qui renvoie si un jeu de cartes a un brelan via un Array
'Si non renvoie 0
'Si oui renvoie la valeur du brelan et les 2 hauteurs

Function brelan(jeu_cartes As Scripting.Dictionary)
    
    Dim isbrelan As Boolean
    Dim cle_brelan As Integer
    Dim hauteur1 As Integer
    Dim hauteur2 As Integer
    Dim Col_valeur As Scripting.Dictionary
    Set Col_valeur = dict_valeur(jeu_cartes)
    
    cle_brelan = 0
    hauteur1 = 0
    hauteur2 = 0
    isbrelan = False
    
    'meilleure brelan
    For Each Key In Col_valeur.Keys
    
        If Col_valeur(Key) = 3 And Key > cle_brelan Then
            cle_brelan = Key
            isbrelan = True
        End If
        
    Next Key
    If cle_brelan = 0 Then GoTo Fin:
    'enlever la cle brelan
    Col_valeur.Remove (cle_brelan)
    
    '1ere meilleure hauteur
    For Each Key In Col_valeur.Keys
    
        If Col_valeur(Key) >= 1 And Key > hauteur1 Then
            hauteur1 = Key
        End If
        
    Next Key
    'enlever la hauteur1
    Col_valeur.Remove (hauteur1)
    
    '2nde meilleure hauteur
    For Each Key In Col_valeur.Keys
    
        If Col_valeur(Key) >= 1 And Key > hauteur2 Then
            hauteur2 = Key
        End If
        
    Next Key
Fin:
    If isbrelan = True Then
        brelan = Array(cle_brelan, hauteur1, hauteur2)
    Else
        brelan = Array(0)
    End If
    
End Function

'Fonction qui renvoie si un jeu de cartes a une double paire via un Array
'Si non renvoie 0
'Si oui renvoie les valeur des paires et la hauteur

Function double_paire(jeu_cartes As Scripting.Dictionary)
    
    Dim ispaire1 As Boolean
    Dim ispaire2 As Boolean
    Dim cle_paire1 As Integer
    Dim cle_paire2 As Integer
    Dim hauteur_ As Integer
    Dim Col_valeur As Scripting.Dictionary
    Set Col_valeur = dict_valeur(jeu_cartes)
    
    cle_paire1 = 0
    cle_paire2 = 0
    ispaire1 = False
    ispaire2 = False
    hauteur_ = 0
    
    'la meilleure 1ere paire
    For Each Key In Col_valeur.Keys
    
        If Col_valeur(Key) = 2 And Key > cle_paire1 Then
            cle_paire1 = Key
            ispaire1 = True
        End If
        
    Next Key
    If cle_paire1 = 0 Then GoTo Fin:
    'enlever la cle paire1
    Col_valeur.Remove (cle_paire1)
    
    'la meilleure 2nde paire
    For Each Key In Col_valeur.Keys
    
        If Col_valeur(Key) = 2 And Key > cle_paire2 Then
            cle_paire2 = Key
            ispaire2 = True
        End If
        
    Next Key
    If cle_paire2 = 0 Then GoTo Fin:
    'enlever la cle brelan
    Col_valeur.Remove (cle_paire2)
    
    'la meilleure hauteur
    For Each Key In Col_valeur.Keys
    
        If Col_valeur(Key) >= 1 And Key > hauteur_ Then
            hauteur_ = Key
        End If
        
    Next Key
Fin:
    If ispaire1 = True And ispaire2 = True Then
        double_paire = Array(cle_paire1, cle_paire2, hauteur_)
    Else
        double_paire = Array(0)
    End If
    
End Function

'Fonction qui renvoie si un jeu de cartes a un paire via un Array
'Si non renvoie 0
'Si oui renvoie la valeur de la paire et les 3 hauteurs

Function paire(jeu_cartes As Scripting.Dictionary)
    
    Dim ispaire As Boolean
    Dim cle_paire As Integer
    Dim hauteur1 As Integer
    Dim hauteur2 As Integer
    Dim hauteur3 As Integer
    
    Dim Col_valeur As Scripting.Dictionary
    Set Col_valeur = dict_valeur(jeu_cartes)
    
    cle_paire = 0
    ispaire = False
    hauteur1 = 0
    hauteur2 = 0
    hauteur3 = 0
    
    'meilleure paire
    For Each Key In Col_valeur.Keys
    
        If Col_valeur(Key) = 2 And Key > cle_paire Then
            cle_paire = Key
            ispaire = True
        End If
        
    Next Key
    If cle_paire = 0 Then GoTo Fin:
    'enlever la cle paire1
    Col_valeur.Remove (cle_paire)
    
    'la 1ere meilleure hauteur
    For Each Key In Col_valeur.Keys
    
        If Col_valeur(Key) >= 1 And Key > hauteur1 Then
            hauteur1 = Key
        End If
        
    Next Key
    
    Col_valeur.Remove (hauteur1)
    
    'la 2nde meilleure hauteur
    For Each Key In Col_valeur.Keys
    
        If Col_valeur(Key) >= 1 And Key > hauteur2 Then
            hauteur2 = Key
        End If
        
    Next Key
    
    Col_valeur.Remove (hauteur2)
    
    'la 3eme meilleure hauteur
    For Each Key In Col_valeur.Keys
    
        If Col_valeur(Key) >= 1 And Key > hauteur3 Then
            hauteur3 = Key
        End If
        
    Next Key
Fin:
    If ispaire = True Then
        paire = Array(cle_paire, hauteur1, hauteur2, hauteur3)
    Else
        paire = Array(0)
    End If
    
End Function

'Fonction qui renvoie les 5meilleures hauteurs de jeu de cartes  via un Array

Function hauteur(jeu_cartes As Scripting.Dictionary)
    
    Dim hauteur_vect As Variant
    ReDim hauteur_vect(1, 5) As Variant
    hauteur_test = 0
    
    For iter = 1 To 5
        
        For Each Key In dict_valeur(jeu_cartes).Keys
    
            If dict_valeur(jeu_cartes)(Key) >= 1 And Key > hauteur_test Then
                hauteur_test = Key
            End If
        
        Next Key
        'enlever la hauteur
        dict_valeur(jeu_cartes).Remove (hauteur_test)
        hauteur_vect(1, iter) = hauteur_test
        hauteur_test = 0
   
    Next iter
   
    hauteur = Array(hauteur_vect(1, 1), hauteur_vect(1, 2), hauteur_vect(1, 3), hauteur_vect(1, 4), hauteur_vect(1, 5))
    
End Function
  
'Cette fonction renvoie un dictionnaire avec des cles allant de 1 à 13 (correspondant aux differentes valeurs d'une cartes)
'et les items sont le nombre d'occurences parmi les cartes de jeu_cartes

Function dict_valeur(jeu_cartes As Scripting.Dictionary) As Scripting.Dictionary

    Dim Collec As New Scripting.Dictionary
    For iterateur = 13 To 1 Step -1
        Collec.Add iterateur, 0
    Next iterateur
     
    For Each Key In jeu_cartes.Keys
        'Ajout d'une occurrence a la cle Key
        Collec(jeu_cartes(Key).get_valeur) = Collec(jeu_cartes(Key).get_valeur) + 1
    Next Key
    
    Set dict_valeur = Collec

End Function

'Cette fonction renvoie un dictionnaire avec des cles allant de 1 à 4 (correspondant aux differentes couleurs d'une cartes)
'et les items sont le nombre d'occurences parmi les cartes de jeu_cartes

Function dict_couleur(jeu_cartes As Scripting.Dictionary) As Scripting.Dictionary
    
    Dim Collec As New Scripting.Dictionary
    Collec.Add "COEUR", 0
    Collec.Add "CARREAU", 0
    Collec.Add "PIQUE", 0
    Collec.Add "TREFLE", 0
    
    For Each Key In jeu_cartes.Keys
        'Ajout d'une occurrence a la cle couleur Key
        Collec(jeu_cartes(Key).get_couleur_string) = Collec(jeu_cartes(Key).get_couleur_string) + 1
    Next Key
    
    Set dict_couleur = Collec
    
End Function


'Cette fonction renvoie un dictionnaire avec des cles allant de 1 à 4 (correspondant aux differentes couleurs d'une cartes)
'et les items sont des dictionnaires
'Ces derniers correspondent à un dict_valeur parmi les cartes de jeu_cartes de couleur = cle

Function dict_valeur_couleur(jeu_cartes As Scripting.Dictionary) As Scripting.Dictionary
    Dim Collec As New Scripting.Dictionary
    
    'Remplissage dict_valeur parmi les cartes de jeu_cartes de couleur = "COEUR"
    Dim dict_val_COEUR As Scripting.Dictionary
    Set dict_val_COEUR = New Scripting.Dictionary
    
    For iterateur = 13 To 1 Step -1
        dict_val_COEUR.Add iterateur, 0
    Next iterateur
    
    For Each Key In jeu_cartes.Keys
        If jeu_cartes(Key).get_couleur_string = "COEUR" Then
            dict_val_COEUR(jeu_cartes(Key).get_valeur) = dict_val_COEUR(jeu_cartes(Key).get_valeur) + 1
        End If
    Next Key
    
    'Remplissage dict_valeur parmi les cartes de jeu_cartes de couleur = "CARREAU"
    Dim dict_val_CARREAU As Scripting.Dictionary
    Set dict_val_CARREAU = New Scripting.Dictionary
    
    For iterateur = 13 To 1 Step -1
        dict_val_CARREAU.Add iterateur, 0
    Next iterateur
    
    For Each Key In jeu_cartes.Keys
        If jeu_cartes(Key).get_couleur_string = "CARREAU" Then
            dict_val_CARREAU(jeu_cartes(Key).get_valeur) = dict_val_CARREAU(jeu_cartes(Key).get_valeur) + 1
        End If
    Next Key
    
    'Remplissage dict_valeur parmi les cartes de jeu_cartes de couleur = "PIQUE"
    Dim dict_val_PIQUE As Scripting.Dictionary
    Set dict_val_PIQUE = New Scripting.Dictionary
    
    For iterateur = 13 To 1 Step -1
        dict_val_PIQUE.Add iterateur, 0
    Next iterateur
    
    For Each Key In jeu_cartes.Keys
        If jeu_cartes(Key).get_couleur_string = "PIQUE" Then
            dict_val_PIQUE(jeu_cartes(Key).get_valeur) = dict_val_PIQUE(jeu_cartes(Key).get_valeur) + 1
        End If
    Next Key
    
    'Remplissage dict_valeur parmi les cartes de jeu_cartes de couleur = "TREFLE"
    Dim dict_val_TREFLE As Scripting.Dictionary
    Set dict_val_TREFLE = New Scripting.Dictionary
    
    For iterateur = 13 To 1 Step -1
        dict_val_TREFLE.Add iterateur, 0
    Next iterateur
    
    For Each Key In jeu_cartes.Keys
        If jeu_cartes(Key).get_couleur_string = "TREFLE" Then
            dict_val_TREFLE(jeu_cartes(Key).get_valeur) = dict_val_TREFLE(jeu_cartes(Key).get_valeur) + 1
        End If
    Next Key
    
    'Remplissage du dictionnaire final
    Collec.Add "COEUR", dict_val_COEUR
    Collec.Add "CARREAU", dict_val_CARREAU
    Collec.Add "PIQUE", dict_val_PIQUE
    Collec.Add "TREFLE", dict_val_TREFLE
    Set dict_valeur_couleur = Collec
    
    
End Function
