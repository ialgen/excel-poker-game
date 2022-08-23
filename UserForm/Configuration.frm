VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Configuration 
   Caption         =   "Configuration de la partie"
   ClientHeight    =   3717
   ClientLeft      =   100
   ClientTop       =   400
   ClientWidth     =   5220
   OleObjectBlob   =   "Configuration.frx":0000
   StartUpPosition =   1  'PropriétaireCentre
End
Attribute VB_Name = "Configuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Base 1

'Cette Userform a pour objectif de recuperer auprès de l'utilisateur les donnees essentielles pour
'initialiser une partie : nombre de participants, la small blind, l'argent initial de chacun des joueurs

'Elle configure aussi l'affichage sur la feuille excel partie en cours

Private Sub UserForm_Initialize()
    BoxJoueurs.AddItem "2"
    BoxJoueurs.AddItem "3"
    BoxJoueurs.AddItem "4"
    BoxJoueurs.AddItem "5"
    BoxJoueurs.AddItem "6"
    BoxJoueurs.Value = "2"
End Sub

Private Sub BoutonLancement_click()

'A/Verification des entrees utilisateur
'1/Blind et argent sont des valeurs numeriques
If Not (IsNumeric(BoxBlind.Value) And IsNumeric(BoxArgent.Value)) Then
    Call erreur("La blind et l' argent des joueurs sont des valeurs numériques.")
    Exit Sub
End If

'2/Blind positive
If BoxBlind.Value < 1 Then
    Call erreur("La valeur de la blind doit être positive.")
    Exit Sub
End If

'3/Stack paye au moins une big blind
If BoxBlind.Value > 0.5 * CLng(BoxArgent.Value) Then 'PBBB
    Debug.Print BoxBlind.Value
    Debug.Print max_blind
    Call erreur("Les joueurs doivent posséder au minimum le double du montant de la blind.")
    Exit Sub
End If

'4/le nombre de participants est un entier entre 2 et 6
If CInt(BoxJoueurs.Value) <> 2 And CInt(BoxJoueurs.Value) <> 3 And CInt(BoxJoueurs.Value) <> 4 And CInt(BoxJoueurs.Value) <> 5 And CInt(BoxJoueurs.Value) <> 6 Then
    Call erreur("Le nombre de participants doit être compris entre 2 et 6.")
    Exit Sub
End If

'idem
If Not (IsNumeric(BoxJoueurs.Value)) Or CInt(BoxJoueurs.Value) < 2 Or CInt(BoxJoueurs.Value) > 6 Then
    Call erreur("La valeur du nombre de joueur doit etre un entier compris entre 2 et 6")
    Exit Sub
End If

'Creation d'une feuille pour la partie
ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count)
Set wsP = ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count)
wsP.name = "Partie en cours"

'B/Affichage sur feuille excel "Partie en cours"

'Background vert
wsP.Cells.Interior.Color = RGB(70, 120, 50)

nb_joueurs = CInt(BoxJoueurs.Value)

'ajustement des colonnes
wsP.Columns(4).ColumnWidth = 4.2
wsP.Columns(5).ColumnWidth = 4.2
wsP.Columns(6).ColumnWidth = 4.8
wsP.Rows.RowHeight = 20

'Ordre joueurs
Set positions = init_positions(CInt(nb_joueurs), 1)
'Affichage des joueurs
For i = 1 To nb_joueurs
    'cellule joueur
    With wsP.Cells(4 * i - 2, 3)
        .name = "Nom_J" & i
        .Value = "Joueur " & i
        .Font.Bold = True
        .Font.name = "Calibri"
        .Font.Size = 11
        .Interior.Color = vbWhite
    End With
    
    'cellules pour la main
    wsP.Cells(4 * i - 2, 4).name = "valeur_carte_1_J" & i
    wsP.Cells(4 * i - 1, 4).name = "couleur_carte_1_J" & i
    wsP.Cells(4 * i - 2, 5).name = "valeur_carte_2_J" & i
    wsP.Cells(4 * i - 1, 5).name = "couleur_carte_2_J" & i
    
    With wsP.Cells(4 * i - 2, 4).Resize(2, 2)
        .Font.name = "Calibri"
        .Font.Size = 12
        .Font.Bold = True
        .Interior.Color = vbWhite
    End With
    
    'cellule position
    wsP.Cells(4 * i - 1, 2).Resize(1, 2).Merge
    With wsP.Cells(4 * i - 1, 2)
        .name = "Position_J" & i
        .Interior.Color = RGB(200, 220, 180)
        .Value = positions(i)
        .Font.Bold = True
        .Font.name = "Calibri"
        .Font.Size = 11
    End With
    
    'cellules stack action et mise
    With wsP.Cells(4 * i - 2, 6).Resize(1, 2)
        .Interior.Color = RGB(200, 220, 180)
        .Font.Bold = True
        .Font.name = "Calibri"
        .Font.Size = 11
    End With
    
    wsP.Cells(4 * i - 2, 6).Value = "Stack"
    With wsP.Cells(4 * i - 2, 7)
        .Value = CLng(BoxArgent.Value)
        .name = "Stack_J" & i
    End With
    
    wsP.Cells(4 * i - 1, 6).Value = "Action"
    wsP.Cells(4 * i - 1, 7).name = "Action_J" & i
    wsP.Cells(4 * i, 6).Value = "Mise"
    wsP.Cells(4 * i, 7).name = "Mise_J" & i
    
    'mises obligatoires
    'relativement complexe car depend du nombre de joueurs (si nb_joueurs=2 les noms ne sont pas les memes)
    If (wsP.Cells(4 * i - 1, 2).Value = "Button / Small Blind" Or wsP.Cells(4 * i - 1, 2).Value = "Small Blind") Then
        wsP.Cells(4 * i, 7).Value = CLng(BoxBlind.Value)
        wsP.Cells(4 * i - 2, 7).Value = wsP.Cells(4 * i - 2, 7).Value - CLng(BoxBlind.Value)
    ElseIf wsP.Cells(4 * i - 1, 2).Value = "Big Blind" Then
        wsP.Cells(4 * i, 7).Value = 2 * CLng(BoxBlind.Value)
        wsP.Cells(4 * i - 2, 7).Value = wsP.Cells(4 * i - 2, 7).Value - 2 * CLng(BoxBlind.Value)
    End If
    
    'intitules et background stack actions et mise
    wsP.Cells(4 * i - 2, 6).Resize(3, 1).Font.Italic = True
       
    With wsP.Cells(4 * i - 1, 6).Resize(2, 2)
        .Interior.Color = RGB(255, 240, 200)
        .Font.name = "Calibri"
        .Font.Size = 11
    End With
    
    wsP.Cells(4 * i - 1, 6).Resize(2, 1).Font.Size = 9
    
    'Centrer les elements dans leurs cellules
    With wsP.Cells(4 * i - 2, 2).Resize(3, 6)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    'bordures
    wsP.Cells(4 * i - 2, 4).Resize(2, 4).Borders.LineStyle = 1
    wsP.Cells(4 * i, 6).Resize(1, 2).Borders.LineStyle = 1
    wsP.Cells(4 * i - 2, 3).Borders.LineStyle = 1
Next i

'Affichage des cartes communes
With wsP.Cells(6, 10).Resize(1, 5)
    .Interior.Color = vbBlack
    .Font.Color = vbWhite
    .Font.Bold = True
    .Font.name = "Calibri"
    .Font.Size = 12
End With

wsP.Cells(6, 10).Resize(1, 3).Merge
wsP.Cells(6, 10).Value = "FLOP"
wsP.Cells(6, 13).Value = "TURN"
wsP.Cells(6, 14).Value = "RIVER"

With wsP.Cells(7, 10).Resize(2, 5)
    .Interior.Color = vbWhite
    .Font.Color = vbBlack
    .Font.Bold = True
    .Font.name = "Calibri"
    .Font.Size = 12
End With

'nom des cellules excel
For j = 1 To 5
    wsP.Cells(7, 9 + j).name = "valeur_tirage_" & j
    wsP.Cells(8, 9 + j).name = "couleur_tirage_" & j
Next j

With wsP.Cells(6, 10).Resize(3, 5)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Borders.LineStyle = 1
End With

'Affichage du pot
With wsP.Cells(11, 12)
    .Interior.Color = vbBlack
    .Font.Color = vbWhite
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Borders.LineStyle = 1
    .Font.name = "Calibri"
    .Font.Size = 12
    .Value = "POT"
End With

'Affichage du pot
With wsP.Cells(12, 12)
    .Interior.Color = vbWhite
    .Font.Color = vbBlack
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Borders.LineStyle = 1
    .Font.name = "Calibri"
    .Font.Size = 12
    .name = "pot"
    .Value = 0
End With


'Page de parametres du jeu
ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets("Partie en cours")
Set wsParametres = ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count)
wsParametres.name = "Parametres"

With wsParametres

    .Cells(1, 1).Value = nb_joueurs
    .Cells(1, 1).name = "Nbre_joueurs"
    .Cells(1, 2).Value = "Nombre de joueurs."
    
    .Cells(2, 1).Value = BoxArgent.Value
    .Cells(2, 1).name = "argent_joueur"
    .Cells(2, 2).Value = "Stack initial par joueur."
    
    .Cells(3, 1).Value = nb_joueurs * BoxArgent.Value
    .Cells(3, 1).name = "argent_en_jeu"
    .Cells(3, 2).Value = "Somme totale des stacks."
    
    .Cells(4, 1).Value = BoxBlind.Value
    .Cells(4, 1).name = "blind"
    .Cells(4, 2).Value = "Valeur de la small blind."
    
    Select Case nb_joueurs
        Case 2
            indice_utg = 1
        Case 3
            indice_utg = 1
        Case Else
            indice_utg = 4
    End Select
    .Cells(5, 1).Value = indice_utg
    .Cells(5, 1).name = "indice_utg"
    .Cells(5, 2).Value = "Indice UTG."
    
    .Cells(6, 1).Value = indice_utg
    .Cells(6, 1).name = "joueur_actif"
    .Cells(6, 2).Value = "Indice joueur actif."
    
    .Cells(7, 1).Value = BoxBlind.Value * 2
    .Cells(7, 1).name = "mise_max"
    .Cells(7, 2).Value = "Valeur de la plus grande mise."
    
    .Cells(8, 1).Value = 0
    .Cells(8, 1).name = "fin_jeu"
    .Cells(8, 2).Value = "Boolean indiquant si la partie est terminee."

End With

Unload Me

End Sub

'Cette fonction renvoie une collection de joueurs ordonnee pour que le premier de parole soit a la position
'DealerPos, elle depend du nombre de joueurs et permet ensuite d'afficher dans la feuille excel "partie en cours"
'le nom de chaque position des joueurs

Function init_positions(nb_joueurs As Integer, DealerPos As Integer) As Collection
    Set init_positions = New Collection

    'Definition des positions
    If nb_joueurs = 2 Then
        init_positions.Add "Button / Small Blind"
        init_positions.Add "Big Blind"
    Else
        init_positions.Add "Button"
        init_positions.Add "Small Blind"
        init_positions.Add "Big Blind"
        init_positions.Add "UTG"
        init_positions.Add "UTG+1"
        init_positions.Add "Cut-Off"

        For i = nb_joueurs + 1 To 6
            init_positions.Remove (nb_joueurs + 1)
        Next i
    End If

    'Positions dans l'ordre
    For j = nb_joueurs To nb_joueurs - DealerPos + 2 Step -1
        init_positions.Add init_positions.item(nb_joueurs), Before:=1
        init_positions.Remove (nb_joueurs + 1)
    Next j

    For i = nb_joueurs + 1 To nb_joueurs
        init_positions.Remove (nb_joueurs + 1)
    Next i
   
End Function

'Renvoie en MsgBox les messages d'erreurs

Private Sub erreur(message As String)
    MsgBox (message)
End Sub

