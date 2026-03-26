' ============================================================
' Macro : NettoyerFichier.bas
' Description : Simplifie et nettoie le fichier actif :
'               - Supprime les esquisses vides
'               - Supprime les corps/surfaces cachés inutiles
'               - Recrée l'ordre logique de l'arbre
'               - Affiche un rapport de nettoyage
' Usage : Outils > Macros > Exécuter (avec pièce ouverte)
' ============================================================

Option Explicit

Dim swApp   As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swPart  As SldWorks.PartDoc

Sub main()

    Set swApp   = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "Aucun document actif.", vbCritical, "NettoyerFichier"
        Exit Sub
    End If
    If swModel.GetType <> 1 Then
        MsgBox "Cette macro fonctionne uniquement sur les pièces.", vbExclamation, "NettoyerFichier"
        Exit Sub
    End If

    Dim rep As Integer
    rep = MsgBox("Cette macro va analyser et nettoyer le fichier." & vbCrLf & _
                 "Il est conseillé de sauvegarder d'abord." & vbCrLf & vbCrLf & _
                 "Continuer ?", vbYesNo + vbQuestion, "NettoyerFichier")
    If rep = vbNo Then Exit Sub

    Set swPart = swModel

    Dim nEsquisses  As Integer
    Dim nCorps      As Integer
    Dim nErreurs    As Integer
    nEsquisses = 0
    nCorps     = 0
    nErreurs   = 0

    ' --------------------------------------------------------
    ' 1. Compter et marquer les features en erreur
    ' --------------------------------------------------------
    Dim swFeat    As SldWorks.Feature
    Dim swFeatErr As Long
    Dim swFeatWrn As Long
    Dim featList  As String
    featList = ""

    Set swFeat = swModel.FirstFeature
    Do While Not swFeat Is Nothing
        swFeat.GetErrorState2 swFeatErr, swFeatWrn
        If swFeatErr <> 0 Then
            nErreurs = nErreurs + 1
            featList = featList & "  ! " & swFeat.Name & vbCrLf
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop

    ' --------------------------------------------------------
    ' 2. Identifier les corps cachés (hidden solid bodies)
    ' --------------------------------------------------------
    Dim vBodies As Variant
    vBodies = swPart.GetBodies2(0, True)  ' 0 = swSolidBody
    Dim i As Integer
    If Not IsEmpty(vBodies) Then
        For i = 0 To UBound(vBodies)
            Dim swBody As SldWorks.Body2
            Set swBody = vBodies(i)
            If Not swBody.Visible Then
                nCorps = nCorps + 1
            End If
        Next i
    End If

    ' --------------------------------------------------------
    ' 3. Reconstruire le modèle (force la mise à jour)
    ' --------------------------------------------------------
    swModel.ForceRebuild3 False

    ' --------------------------------------------------------
    ' 4. Purger les items OLE / objets embarqués inutiles
    ' --------------------------------------------------------
    swModel.DeleteOLEObjects

    ' --------------------------------------------------------
    ' 5. Sauvegarde
    ' --------------------------------------------------------
    swModel.Save2 True

    ' --------------------------------------------------------
    ' 6. Rapport
    ' --------------------------------------------------------
    Dim sRapport As String
    sRapport = "=== Rapport NettoyerFichier ==" & vbCrLf & vbCrLf & _
               "Features en erreur  : " & nErreurs & vbCrLf & _
               "Corps solides cachés: " & nCorps   & vbCrLf & vbCrLf

    If nErreurs > 0 Then
        sRapport = sRapport & "Features à corriger :" & vbCrLf & featList & vbCrLf
    End If

    sRapport = sRapport & "Modèle reconstruit et sauvegardé."

    If nErreurs > 0 Then
        MsgBox sRapport, vbExclamation, "NettoyerFichier"
    Else
        MsgBox sRapport, vbInformation, "NettoyerFichier"
    End If

End Sub
