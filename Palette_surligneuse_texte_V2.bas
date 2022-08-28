Attribute VB_Name = "Palette_surligneuse_texte_V2"
''''' constantes globales '''''
Public Const nbr_base_btn As Integer = 9 'min 1
Public Const taille_btn As Integer = 3 'min 1
Public Const espacement_btn As Integer = 0 'min 0
Public Const ligne_de_base_btn As Integer = 5 'min 1
''''' en les modifiants vous pouvez quelque peu personnaliser l'outil '''''

Sub main()
    Dim oClasseur: Set oClasseur = ActiveWorkbook
    Dim oFeuille: Set oFeuille = oClasseur.ActiveSheet
    Dim btn As Button
    Dim colonne_btn As Integer: colonne_btn = 1 'min 1
    Dim id_btn As Integer: id_btn = 1
    Dim cpt_suppression As Integer
    
    ' checker la préexistance des boutons
    ' tout supprimer est plus simple
    For Each btn_presents In oFeuille.Buttons
        Dim cpt_verif As Integer
        For cpt_verif = 1 To nbr_base_btn Step 1
            On Error GoTo pblm_dacces_nom
            If InStr(btn_presents.Name, nom_btn_parametrer(cpt_verif)) = 1 Or _
            InStr(btn_presents.Name, nom_btn_parametrer(-1)) = 1 Then
                btn_presents.Delete
                Exit For
            End If
pblm_dacces_nom:
        Next cpt_verif
    Next
    
    ' tout créer
    For i = ligne_de_base_btn To ligne_de_base_btn + (nbr_base_btn - 1) * _
    (espacement_btn + taille_btn) Step (espacement_btn + taille_btn)
      Set plage = oFeuille.Range(Cells(i, colonne_btn), Cells(i, colonne_btn))
      Set btn = oFeuille.Buttons.Add(plage.Left, plage.Top, plage.Width, _
      plage.Height * taille_btn)
      
      With btn
        .OnAction = "colorier_text_par_bouton"
        .Caption = cap_btn_parametrer(id_btn)
        .Name = nom_btn_parametrer2(id_btn, id_btn)
      End With
      
      id_btn = id_btn + 1
    Next i
End Sub

Public Function nom_btn_parametrer(id As Integer) As String
    Select Case id
        Case 1
            nom_btn_parametrer = "btn_Rouge"
        Case 2
            nom_btn_parametrer = "btn_Noir"
        Case 3
            nom_btn_parametrer = "btn_Vert"
        Case 4
            nom_btn_parametrer = "btn_surligne_Rouge"
        Case 5
            nom_btn_parametrer = "btn_surligne_Blanc"
        Case 6
            nom_btn_parametrer = "btn_surligne_Jaune"
        Case 7
            nom_btn_parametrer = "btn_Gras"
        Case 8
            nom_btn_parametrer = "btn_Souligne"
        Case 9
            nom_btn_parametrer = "btn_normal"
        Case Else
            nom_btn_parametrer = "btn_Rien"
    End Select
End Function

Public Function nom_btn_parametrer2(id As Integer, nbr As Integer) As String
    Select Case id
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9
            nom_btn_parametrer2 = nom_btn_parametrer(id)
        Case Else
            nom_btn_parametrer2 = "btn_Rien" & nbr
    End Select
End Function

Public Function cap_btn_parametrer(id As Integer) As String
    Select Case id
        Case 1
            cap_btn_parametrer = "Rouge :" & Chr(13) & Chr(10) & "URGENT"
        Case 2
            cap_btn_parametrer = "Noir :" & Chr(13) & Chr(10) & "EN COURS"
        Case 3
            cap_btn_parametrer = "Vert :" & Chr(13) & Chr(10) & "VALIDÉ"
        Case 4
            cap_btn_parametrer = "Surligner :" & Chr(13) & Chr(10) & "ROUGE"
        Case 5
            cap_btn_parametrer = "Surligner :" & Chr(13) & Chr(10) & "BLANC"
        Case 6
            cap_btn_parametrer = "Surligner :" & Chr(13) & Chr(10) & "JAUNE"
        Case 7
            cap_btn_parametrer = "Gras"
        Case 8
            cap_btn_parametrer = "Souligné"
        Case 9
            cap_btn_parametrer = "Normal"
        Case Else
            cap_btn_parametrer = "Rien :" & Chr(13) & Chr(10) & "ÉRREUR"
    End Select
End Function

Function colorier_text_par_bouton()

    Dim oCellule: Set oCellule = ActiveCell
    Dim ButtonText As String
    
    ButtonText = Application.Caller
    
    Select Case ButtonText
        Case nom_btn_parametrer(1)
            With oCellule.Font
                .Color = -16776961
                .TintAndShade = 0
            End With
        Case nom_btn_parametrer(2)
            With oCellule.Font
                .Color = xlThemeColorLight1
                .TintAndShade = 0
            End With
        Case nom_btn_parametrer(3)
            With oCellule.Font
                .Color = -11480942
                .TintAndShade = 0
            End With
        Case nom_btn_parametrer(4)
            With oCellule.Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
            With oCellule.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Case nom_btn_parametrer(5)
            With oCellule.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With oCellule.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Case nom_btn_parametrer(6)
            With oCellule.Font
                .Color = xlThemeColorDark1
                .TintAndShade = 0
            End With
            With oCellule.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Case nom_btn_parametrer(7)
            oCellule.Font.Bold = True
        Case nom_btn_parametrer(8)
            oCellule.Font.Underline = xlUnderlineStyleSingle
        Case nom_btn_parametrer(9)
            oCellule.Font.Bold = False
            oCellule.Font.Underline = xlUnderlineStyleNone
        Case Else
            With oCellule.Font
                .Color = xlAutomatic
                .TintAndShade = 0
            End With
    End Select
    
End Function

Function nombre_en_lettre(num As Integer) As String

    nombre_en_lettre = Split(Cells(1, num).Address, "$")(1)
    
End Function

Sub setup_first_use()
    
    Dim texte_avertissement As String: texte_avertissement = _
        "ATTENTION : Cette macro ainsi que son fonctionnement" & Chr(13) & Chr(10) & _
        "peuvent potentiellement entraîner des problèmes non découverts" & Chr(13) & Chr(10) & _
        "lors d'une utilisation en relation avec différents macros et codes" & Chr(13) & Chr(10) & _
        "dans un même fichier." & Chr(13) & Chr(10) & _
        "Par mesure de sécurité, veillez bien utiliser cet outil dans le cadre de" & Chr(13) & Chr(10) & _
        "prise de note ou sinon en toute conscience de cause." & Chr(13) & Chr(10) & _
        "Merci pour votre confiance. Cordialement." & Chr(13) & Chr(10) & _
        "               Corentin Le Goff, Concepteur et Développeur de cette macro."
    Dim msgbox_avertissement
    msgbox_avertissement = MsgBox(texte_avertissement, vbYesNo, "Message du développeur : Corentin Le Goff")
    If msgbox_avertissement = vbNo Then
        MsgBox prompt:="J'espère que vous utiliserez cet outil sur un classeur qui ne craint rien" & Chr(13) & Chr(10) & _
        "               Corentin Le Goff, Concepteur et Développeur de cette macro.", _
        Title:="Message du développeur : Corentin Le Goff"
        Exit Sub
    End If
    
    Dim oClasseur: Set oClasseur = ActiveWorkbook
    Dim nom_feuille_tache As String: nom_feuille_tache = "Taches"
    On Error GoTo feuille_pre_existante
    Sheets.Add.Name = nom_feuille_tache
    Dim oFeuille: Set oFeuille = Sheets(nom_feuille_tache)
    Dim btn As Button
    Dim texte_de_finalisation As String: texte_de_finalisation = _
    "Finalisation : Veuillez ajouter le code de la cellule B2 de [Tache] " & Chr(13) & Chr(10) & _
        "dans la feuille renommée [Tache] après avoir fait Alt+F11" & Chr(13) & Chr(10) & _
        "PS : les consignes ci-dessous vous seront rappelées" & Chr(13) & Chr(10) & _
        " -enlevez après le collage du code les quotes devant [Private] au début," & Chr(13) & Chr(10) & _
        " -et après [End Sub] à la fin ligne 32" & Chr(13) & Chr(10) & _
        "!!! Il arrive qu'un [End Sub] s'ajoute ligne 3 : supprimez le !!!"
        
    Set btn = oFeuille.Buttons.Add(0, 0, 60, 4 * 14.5)
    With btn
        .OnAction = "main"
        .Caption = "ÉTAPE 3 : Cliquez ici"
        .Name = "btn_main"
    End With
    MsgBox prompt:=texte_de_finalisation, Title:="Finalisation"
    oClasseur.Sheets(nom_feuille_tache).Cells(2, 2).Value = _
        "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & Chr(13) & Chr(10) & _
        "    On Error GoTo Fin_de_tache" & Chr(13) & Chr(10) & _
        "    Dim numrow As Integer: numrow = ActiveWindow.VisibleRange.Row" & Chr(13) & Chr(10) & _
        "    Dim numcol As Integer: numcol = ActiveWindow.VisibleRange.Column" & Chr(13) & Chr(10) & _
        "    Dim oClasseur: Set oClasseur = ActiveWorkbook" & Chr(13) & Chr(10) & _
        "    Dim oFeuille: Set oFeuille = oClasseur.ActiveSheet" & Chr(13) & Chr(10) & _
        "    Dim rng As Range: Set rng = ActiveSheet.Range(nombre_en_lettre(numcol) " & _
                "& numrow + ligne_de_base_btn - 1)" & Chr(13) & Chr(10) & _
        "    Dim num_shape_feuille As Integer" & Chr(13) & Chr(10) & _
        "    Dim num_nom_a_tester As Integer" & Chr(13) & Chr(10) & _
        "    Dim ordre As Integer: ordre = 1" & Chr(13) & Chr(10) & _
        " " & Chr(13) & Chr(10) & _
        "    For num_shape_feuille = 1 To oFeuille.Shapes.Count Step 1" & Chr(13) & Chr(10) & _
        "        For num_nom_a_tester = 1 To nbr_base_btn Step 1" & Chr(13) & Chr(10) & _
        "            If InStr(oFeuille.Shapes(num_shape_feuille).Name, nom_btn_parametrer(num_nom_a_tester)) = 1 Or _" & Chr(13) & Chr(10) & _
        "            InStr(oFeuille.Shapes(num_shape_feuille).Name, nom_btn_parametrer(-1)) = 1 Then" & Chr(13) & Chr(10) & _
        "                With oFeuille.Shapes(nom_btn_parametrer(num_nom_a_tester))" & Chr(13) & Chr(10) & _
        "                    .Top = rng.Top + (ordre - 1) * rng.Height * taille_btn" & Chr(13) & Chr(10) & _
        "                    .Left = rng.Left" & Chr(13) & Chr(10)
    oClasseur.Sheets(nom_feuille_tache).Cells(2, 2).Value = _
    oClasseur.Sheets(nom_feuille_tache).Cells(2, 2).Value & _
        "                    .Width = rng.Width" & Chr(13) & Chr(10) & _
        "                    .Height = rng.Height * taille_btn" & Chr(13) & Chr(10) & _
        "                End With" & Chr(13) & Chr(10) & _
        "                ordre = ordre + 1" & Chr(13) & Chr(10) & _
        "                Exit For" & Chr(13) & Chr(10) & _
        "            End If" & Chr(13) & Chr(10) & _
        "        Next num_nom_a_tester" & Chr(13) & Chr(10) & _
        "    Next num_shape_feuille" & Chr(13) & Chr(10) & _
        "    Exit Sub" & Chr(13) & Chr(10) & _
        " " & Chr(13) & Chr(10) & _
        "Fin_de_tache:" & Chr(13) & Chr(10) & _
        "End Sub"
    oClasseur.Sheets(nom_feuille_tache).Cells(2, 2).RowHeight = _
    oClasseur.Sheets(nom_feuille_tache).Cells(1, 2).RowHeight
    
    oClasseur.Sheets(nom_feuille_tache).Cells(3, 2).Value = _
        "ÉTAPE 2 :" & Chr(13) & Chr(10) & texte_de_finalisation
    oClasseur.Sheets(nom_feuille_tache).Cells(3, 2).RowHeight = _
    oClasseur.Sheets(nom_feuille_tache).Cells(1, 2).RowHeight
    
    With oClasseur.Sheets(nom_feuille_tache).Cells(3, 2).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With oClasseur.Sheets(nom_feuille_tache).Cells(3, 2).Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
    
    Exit Sub
    
    '''''''''''''''''''''''''''''''''''''''''''''
    ' Il faudra ajouter dans la feuille "Taches"'
    ' la sub ci-dessous précédée des ''         '
    '''''''''''''''''''''''''''''''''''''''''''''
    
''Private Sub Worksheet_SelectionChange(ByVal Target As Range)
''    On Error GoTo Fin_de_tache
''    Dim numrow As Integer: numrow = ActiveWindow.VisibleRange.Row
''    Dim numcol As Integer: numcol = ActiveWindow.VisibleRange.Column
''    Dim oClasseur: Set oClasseur = ActiveWorkbook
''    Dim oFeuille: Set oFeuille = oClasseur.ActiveSheet
''    Dim rng As Range: Set rng = ActiveSheet.Range(nombre_en_lettre(numcol) & numrow + ligne_de_base_btn - 1)
''    Dim num_shape_feuille As Integer
''    Dim num_nom_a_tester As Integer
''    Dim ordre As Integer: ordre = 1
''
''    For num_shape_feuille = 1 To oFeuille.Shapes.Count Step 1
''        For num_nom_a_tester = 1 To nbr_base_btn Step 1
''            If InStr(oFeuille.Shapes(num_shape_feuille).Name, nom_btn_parametrer(num_nom_a_tester)) = 1 Or _
''            InStr(oFeuille.Shapes(num_shape_feuille).Name, nom_btn_parametrer(-1)) = 1 Then
''                With oFeuille.Shapes(nom_btn_parametrer(num_nom_a_tester))
''                    .Top = rng.Top + (ordre - 1) * rng.Height * taille_btn
''                    .Left = rng.Left
''                    .Width = rng.Width
''                    .Height = rng.Height * taille_btn
''                End With
''                ordre = ordre + 1
''                Exit For
''            End If
''        Next num_nom_a_tester
''    Next num_shape_feuille
''    Exit Sub
''
''Fin_de_tache:
''End Sub
    
feuille_pre_existante:
    MsgBox prompt:="Erreur : La feuille [Tache] existe déjà." & Chr(13) & Chr(10) & _
    "A la fermeture de cette fenêtre, une autre vous proposera de surpprimer la feuille vierge créée" _
    , Title:="Erreur : Feuille préexistante"
    ActiveSheet.Delete
    End Sub
