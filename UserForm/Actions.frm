VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Actions 
   Caption         =   "Actions"
   ClientHeight    =   3535
   ClientLeft      =   100
   ClientTop       =   400
   ClientWidth     =   5200
   OleObjectBlob   =   "Actions.frx":0000
   StartUpPosition =   1  'PropriétaireCentre
End
Attribute VB_Name = "Actions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Base 1

'Cette UserForm demande l'action du joueur a ce stade de la partie
'son action (check , passer ...)
'si il y a une mise/ on recuperer les valeurs
'On renvoie les infos correspondantes sur les feuilles excel

Private Sub UserForm_Initialize()

'Intialisations de variables
Dim wsP As Worksheet
Set wsP = ActiveWorkbook.Worksheets("Partie en cours")
Numero_J_Parole = ThisWorkbook.Worksheets("Parametres").Range("joueur_actif").Value
MiseMax = ThisWorkbook.Worksheets("Parametres").Range("mise_max").Value

'Conditions sur les actions suivant la situation du joueur dans le jeu
If (wsP.Range("Mise_J" & Numero_J_Parole).Value < MiseMax Or _
    wsP.Range("Mise_J" & Numero_J_Parole) = Null) And _
    wsP.Range("Stack_J" & Numero_J_Parole).Value <> 0 Then

    Box_Action.AddItem "passe"
    Box_Action.AddItem "suis"
    Box_Action.AddItem "relance"
    Box_Action.Value = "passe"

ElseIf wsP.Range("Stack_J" & Numero_J_Parole).Value = 0 Then
    Box_Action.AddItem "check"
    Box_Action.Value = "check"
Else
    Box_Action.AddItem "check"
    Box_Action.AddItem "mise"
    Box_Action.Value = "check"
End If

End Sub

Private Sub Cmd_Valider_Click()

'Initialisation de variables
Numero_J_Parole = ThisWorkbook.Worksheets("Parametres").Range("joueur_actif").Value
MiseMax = ThisWorkbook.Worksheets("Parametres").Range("mise_max").Value

Dim wsP As Worksheet
Set wsP = ActiveWorkbook.Worksheets("Partie en cours")
Set ws_param = ActiveWorkbook.Worksheets("Parametres")
If TxtBox_mise.Value = "" Then
    TxtBox_mise.Value = 0
End If

'Verifications de certaines conditions sur les entrees utilisateurs

'1/selection de l'action que parmi les propositions
If (Box_Action.Value <> "passe") And (Box_Action.Value <> "suis") And (Box_Action.Value <> "relance") And (Box_Action.Value <> "mise") And (Box_Action.Value <> "check") Then
    Call erreur("L' action du joueur doit faire partie des propositions.")
    Unload Me ' on enleve l'ancien message pour mieux le rouvrir
    Actions.Show
    Exit Sub
End If

'2/Si mise, il faut remplir la mise avec une valeur numérique
If Not (IsNumeric(TxtBox_mise.Value)) And (Box_Action.Value = "mise" Or Box_Action.Value = "relance") Then
    Call erreur("La mise du joueur doit être une valeur numérique.")
    Unload Me ' on enleve l'ancien message pour mieux le rouvrir
    Actions.Show
    Exit Sub
End If

'3/Si mise, il faut qu'elle soit plus grande que big blind
If TxtBox_mise.Value < (MiseMax - wsP.Range("Mise_J" & Numero_J_Parole).Value) And (Box_Action.Value = "mise" Or Box_Action.Value = "relance") Then
    Call erreur("La valeur de la mise doit être supérieure à mise la plus haute.")
    Unload Me ' on enleve l'ancien message pour mieux le rouvrir
    Actions.Show
    Exit Sub
End If

'4/Si mise, il faut qu'elle soit plus grande que la mise maximum du jeu
If (CInt(TxtBox_mise.Value) < MiseMax) And (Box_Action.Value = "mise" Or Box_Action.Value = "relance") Then
    Call erreur("La valeur de la mise doit être supérieure à mise la plus haute.")
    Unload Me ' on enleve l'ancien message pour mieux le rouvrir
    Actions.Show
    Exit Sub
End If

'4/Si mise, il faut qu'elle soit plus inferieure au stack
If (TxtBox_mise.Value - wsP.Range("Mise_J" & Numero_J_Parole).Value) > wsP.Range("Stack_J" & Numero_J_Parole).Value And (Box_Action.Value = "mise" Or Box_Action.Value = "relance") Then
    Call erreur("La valeur de la mise ne peut dépasser celle du stack.")
    Unload Me
    Actions.Show
    Exit Sub
End If

'Remplissage des feuilles excel
Select Case Box_Action.Value
    Case Is = "passe"
        wsP.Range("Action_J" & Numero_J_Parole).Value = "passe"

    Case Is = "suis"
        With wsP
            .Range("Action_J" & Numero_J_Parole).Value = "suis"
            If .Range("Stack_J" & Numero_J_Parole).Value + .Range("Mise_J" & Numero_J_Parole).Value - MiseMax <= 0 Then
                .Range("Mise_J" & Numero_J_Parole).Value = .Range("Mise_J" & Numero_J_Parole).Value + .Range("Stack_J" & Numero_J_Parole).Value
                .Range("Stack_J" & Numero_J_Parole).Value = 0
            Else
                .Range("Stack_J" & Numero_J_Parole).Value = .Range("Stack_J" & Numero_J_Parole).Value + .Range("Mise_J" & Numero_J_Parole).Value - MiseMax
                .Range("Mise_J" & Numero_J_Parole).Value = MiseMax
            End If
        End With

    Case Is = "check"
        wsP.Range("Action_J" & Numero_J_Parole).Value = "check"

    Case Is = "mise"
        wsP.Range("Action_J" & Numero_J_Parole).Value = "mise"
        wsP.Range("Stack_J" & Numero_J_Parole).Value = wsP.Range("Stack_J" & Numero_J_Parole).Value + wsP.Range("Mise_J" & Numero_J_Parole).Value - TxtBox_mise.Value
        wsP.Range("Mise_J" & Numero_J_Parole).Value = TxtBox_mise.Value
        ws_param.Range("mise_max").Value = TxtBox_mise.Value

    Case Is = "relance"
        wsP.Range("Action_J" & Numero_J_Parole).Value = "relance"
        wsP.Range("Stack_J" & Numero_J_Parole).Value = wsP.Range("Stack_J" & Numero_J_Parole).Value + wsP.Range("Mise_J" & Numero_J_Parole).Value - TxtBox_mise.Value
        wsP.Range("Mise_J" & Numero_J_Parole).Value = TxtBox_mise.Value
        ws_param.Range("mise_max").Value = TxtBox_mise.Value

End Select

Unload Me

End Sub
'Renvoie en MsgBox les messages d'erreurs
Private Sub erreur(message As String)
    MsgBox (message)
End Sub
