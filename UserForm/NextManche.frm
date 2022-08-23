VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NextManche 
   Caption         =   "Manche Suivante"
   ClientHeight    =   2359
   ClientLeft      =   100
   ClientTop       =   400
   ClientWidth     =   5080
   OleObjectBlob   =   "NextManche.frx":0000
   StartUpPosition =   1  'PropriétaireCentre
End
Attribute VB_Name = "NextManche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Cette UserForm sert seulement a sortir prématurement de la partie
'cette proposition est faite à la fin de chaque manche si on clique sur Quitter

Private Sub Cmd_Continuer_Click()
Unload Me
End Sub

Private Sub Cmd_Quitter_Click()
fin_du_jeu = 0
If MsgBox("Etes - vous sûr de vouloir arrêter la partie ? ", vbYesNo) = vbYes Then
    ThisWorkbook.Worksheets("Parametres").Range("fin_jeu").Value = 1
End If
Unload Me

End Sub
