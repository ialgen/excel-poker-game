Attribute VB_Name = "main"
Option Explicit
Option Base 1

Sub une_partie()

'Variables d'interface
Dim ws_table As Worksheet
Dim ws_param As Worksheet
Dim ws_lancement As Worksheet

'Variables generales
Dim combi As Combinaisons

'Variables du jeu
Dim un_jeu As jeu
Dim joueur_ As joueur
Dim argent_total As Long
Dim nb_joueurs, nbj, last_reveal, indice_utg, new_indice_utg, gagnant As Integer

'Variables temporaires
Dim i, j, k, iter, manche_i, num_joueur_actif, num_col_joueur_actif As Integer
Dim une_main As hand
Dim une_carte As carte
Dim ensemble_cartes As Scripting.Dictionary
Dim rang_main As Variant
Dim resultats_manche As Collection
Dim positions As Collection
Dim meme_hauteur, changement_phase As Boolean
Dim MsgStr As String

Configuration.Show

Set ws_table = ThisWorkbook.Worksheets("Partie en cours")
Set ws_param = ThisWorkbook.Worksheets("Parametres")
Set ws_lancement = ThisWorkbook.Worksheets("lancement_partie")

Set combi = New Combinaisons

'initialisation du jeu
Set un_jeu = New jeu
Call initialisation_class_jeu(un_jeu, ws_table, ws_param)

'initialisation le nombre de participants a cette manche
nb_joueurs = un_jeu.get_joueurs.count

'initialiser get joueur en jeu
ws_table.Activate
For manche_i = 1 To 10

    'Melange du paquet
    Call un_jeu.melange

    'Distribution des cartes
    Call distribution_cartes(un_jeu)

    'Initialisation des variables de la manche
    Set ensemble_cartes = New Scripting.Dictionary
    changement_phase = True
    indice_utg = CInt(ws_param.Range("indice_utg").Value)
    ws_param.Range("mise_max").Value = 2 * CInt(un_jeu.blind)
    
    'Tour de mise de main
    While test_de_mise(un_jeu, ws_param.Range("mise_max").Value, changement_phase) = False
        nbj = nb_joueurs_en_manche(un_jeu)
        If nbj = 1 Then GoTo Fin_manche
        
            'Actions de tous les joueurs
            Call phase_actions(un_jeu, ws_table, ws_param, ws_lancement)
        
        changement_phase = False
    Wend

    'Actualisation du pot
    Call actualisation_pot(un_jeu, ws_table)
    
    'Reveal du flop
    Call reveal_cartes_communes(un_jeu, ensemble_cartes, ws_table, ws_lancement, 1, 3)
    last_reveal = 3

    'Tour de mise de flop
    changement_phase = True
    While test_de_mise(un_jeu, ws_param.Range("mise_max"), changement_phase) = False
         
        nbj = nb_joueurs_en_manche(un_jeu)
        If nbj = 1 Then GoTo Fin_manche
        
            'Actions de tous les joueurs
            Call phase_actions(un_jeu, ws_table, ws_param, ws_lancement)
            
            
        changement_phase = False
    Wend
    
    'Actualisation du pot
    Call actualisation_pot(un_jeu, ws_table)
    
    'Reveal du turn
    Call reveal_cartes_communes(un_jeu, ensemble_cartes, ws_table, ws_lancement, 4, 4)
    last_reveal = 4

    'Tour de mise de turn
    changement_phase = True
    While test_de_mise(un_jeu, ws_param.Range("mise_max"), changement_phase) = False
    
        nbj = nb_joueurs_en_manche(un_jeu)
        If nbj = 1 Then GoTo Fin_manche
        
            'Actions de tous les joueurs
            Call phase_actions(un_jeu, ws_table, ws_param, ws_lancement)
            
        changement_phase = False
    Wend
    
    'Actualisation du pot
    Call actualisation_pot(un_jeu, ws_table)

    'Reveal de la river
    Call reveal_cartes_communes(un_jeu, ensemble_cartes, ws_table, ws_lancement, 5, 5)
    last_reveal = 5
   
    'Tour de mise de river
    changement_phase = True
    While test_de_mise(un_jeu, ws_param.Range("mise_max"), changement_phase) = False
        nbj = nb_joueurs_en_manche(un_jeu)
        If nbj = 1 Then GoTo Fin_manche
         
            'Actions de tous les joueurs
            Call phase_actions(un_jeu, ws_table, ws_param, ws_lancement)
            
        changement_phase = False
    Wend
    
    'Actualisation du pot
    Call actualisation_pot(un_jeu, ws_table)

Fin_manche:
    'Reveal de toutes les cartes
    Call reveal_all(un_jeu, ws_table, ws_lancement)
    If last_reveal < 5 Then
        Call reveal_cartes_communes(un_jeu, ensemble_cartes, ws_table, ws_lancement, last_reveal + 1, 5)
    End If
    
    'Calcul des resultats de la manche
    ensemble_cartes.Add 6, une_carte
    ensemble_cartes.Add 7, une_carte

    Set resultats_manche = New Collection
    Set resultats_manche = resultats(un_jeu, ensemble_cartes)
    
    'Affichage gagnant de cette manche
    MsgStr = "Le joueur " & resultats_manche.item(1)(3) & " a gagnŽ avec : " & vbCrLf & _
            "Combinaison : " & combi.type_best_main(CInt(resultats_manche.item(1)(1)))
            
    For iter = 1 To UBound(resultats_manche.item(1)(2))
        MsgStr = MsgStr + vbCrLf & "Hauteur " & iter & " : " & hauteur_string(CInt(resultats_manche.item(1)(2)(iter)))
    Next iter
    MsgBox (MsgStr)
    
    'Actualisation de la variable un_jeu avec les gains des joueurs
    Call actualisation_jeu(ws_table, ws_param, un_jeu, resultats_manche)
    
    'Actualisation de l'affichage du jeu pour une nouvelle manche
    Call actualisation_affichage(ws_table, un_jeu)
    ws_param.Range("mise_max") = 2 * CInt(un_jeu.blind)

    'mise a jour des joueurs restants
    Call mise_a_jour_joueurs(un_jeu)
    nb_joueurs = un_jeu.get_joueurs_en_jeu.count

    'les Sorties potentielles
        '1/Fin du jeu
    If nb_joueurs = 1 Then
        MsgBox "BRAVO ! Le joueur " & un_jeu.get_joueurs_en_jeu.item(1).get_number & _
                " est le grand gagnant " & vbCrLf & _
                "avec un stack de " & un_jeu.get_joueurs_en_jeu.item(1).get_stack, _
                vbOKOnly + vbInformation
        Exit Sub
    End If

        '2/Sortie prematurŽe ?
    NextManche.Show
    If ws_param.Range("fin_jeu").Value = 1 Then Exit Sub
        
    'Ordre joueurs pour nouvelle manche
    If CInt(ws_param.Range("indice_utg")) = CInt(nb_joueurs) Then
        new_indice_utg = 1
    Else
        new_indice_utg = CInt(ws_param.Range("indice_utg")) + 1
    End If

    ws_param.Range("indice_utg").Value = new_indice_utg
    Set positions = reinit_positions(CInt(nb_joueurs), CInt(new_indice_utg))
    
    'Affichage des positions des joueurs et leurs blinds potentielles
    Call affichage_pos_blind(un_jeu, positions, ws_table, ws_param)
            
Next manche_i

gagnant = 1
For Each joueur_ In un_jeu.get_joueurs
    If joueur_.get_stack > un_jeu.get_joueurs.item(gagnant).get_stack Then
        gagnant = joueur_.get_number
    End If
Next

MsgBox "BRAVO ! Le joueur " & gagnant & _
        " est le grand gagnant " & vbCrLf & _
        "avec un stack de " & un_jeu.get_joueurs_en_jeu.item(gagnant).get_stack, _
        vbOKOnly + vbInformation
Exit Sub

End Sub












