Attribute VB_Name = "nettoyage"
Option Base 1

Sub nettoyage()

'Suppression des feuilles de partie.

Application.DisplayAlerts = False
ThisWorkbook.Worksheets("Partie en cours").Delete
ThisWorkbook.Worksheets("Parametres").Delete
Application.DisplayAlerts = True

End Sub




