Sub creation_sommaire()
    
    If Worksheets(1).Name = "Sommaire" Then ' Si la première feuille s'appelle "Sommaire"
        Application.DisplayAlerts = False   ' On désactive les alertes
        Worksheets(1).Delete                ' On supprime la feuille
        Application.DisplayAlerts = True    ' On re-active les alertes
    End If                                  ' Fin de la condition
      
    Sheets.Add before:=Worksheets(1)        ' On crée une feuille en position 1
    ActiveSheet.Name = "Sommaire"           ' On la renomme "Sommaire"
    
            ' Dans la cellule A1 on écrit:
    [a1] = "     SOMMAIRE    "
    
    Dim ligne As Integer
    ligne = 1   ' Création d'une variable qui sert à gérer le n° de la 1ère ligne
    
    Dim feuille As Worksheet
    For Each feuille In Worksheets          ' On récupère chaque feuilles et on lui associe un lien hypertexte
        If feuille.Name <> "Sommaire" Then  ' Si la feuille est différente de 'Sommaire'
        ActiveSheet.Hyperlinks.Add Anchor:=Cells(ligne, 1), Address:="", SubAddress:="'" & feuille.Name & "'!A1", TextToDisplay:=feuille.Name
        End If

                ' Modification de l'apparence des cellules du sommaire
        With Cells(ligne, 1)
            .Font.Size = 16                 ' Taille de la police en points
            .HorizontalAlignment = xlCenter ' Alignement centré
            .Borders.LineStyle = xlContinuous ' Style de bordure continu
            
                ' Epaisseur des bordures
            '.Borders.Weight = xlThin ' Très fin
            '.Borders.Weight = xlThin ' Fin
            '.Borders.Weight = xlThin ' Moyen
            .Borders.Weight = xlThin ' Epais
                ' Fin des choix d'épaisseurs
            
            .Font.Bold = True               ' Texte en gras
            .Font.Color = RGB(0, 0, 250)    ' Couleur du texte en RGB
            .Interior.Color = RGB(240, 225, 225) ' Couleur de fond RGB
        End With
        
        ligne = ligne + 1               ' On passe à la feuille suivante
    Next
    
    ' Modification de la cellule A1
    With Cells(1, 1)
        .Font.Size = 18
        .Font.Color = RGB(0, 0, 0)
        .Interior.Color = RGB(200, 240, 255)
        .EntireRow.AutoFit              ' Ajuste automatiquement la taille de la ligne
        .EntireColumn.AutoFit           ' Ajuste automatiquement la taille de la colonne
    End With
    
        ' Défini où ajouter le bouton
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sommaire")
      
    ' Ajoute le bouton
    Dim btn As Button
    Set btn = ws.Buttons.Add(150, 15, 100, 50) ' Modifie les coordonnées et la taille du bouton
    With btn
        .Name = "BoutonMacroSommaire"
        .Caption = "Mettre à jour le Sommaire"
        .OnAction = "creation_sommaire" ' Nom de la macro à exécuter
    End With
    
End Sub
