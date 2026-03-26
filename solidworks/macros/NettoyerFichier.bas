' ==============================================================
' Macro : NettoyerFichier.bas
' Description : Diagnostique et nettoie le fichier actif
' Usage : Outils > Macros > Executer (avec piece ouverte)
' ==============================================================

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
        MsgBox "Cette macro fonctionne uniquement sur les pieces.", vbExclamation, "NettoyerFichier"
        Exit Sub
    End If

    Dim rep As Integer
    rep = MsgBox("Cette macro va analyser et nettoyer le fichier." & vbCrLf & "Sauvegardez d'abord. Continuer ?", vbYesNo + vbQuestion, "NettoyerFichier")
    If rep = vbNo Then Exit Sub

    Set swPart = swModel

    Dim nErreurs As Integer
    Dim nCorps   As Integer
    Dim featList As String
    nErreurs = 0
    nCorps   = 0
    featList = ""

    Dim swFeat    As SldWorks.Feature
    Dim swFeatErr As Long
    Dim swFeatWrn As Long

    Set swFeat = swModel.FirstFeature
    Do While Not swFeat Is Nothing
        swFeat.GetErrorState2 swFeatErr, swFeatWrn
        If swFeatErr <> 0 Then
            nErreurs = nErreurs + 1
            featList = featList & "  ! " & swFeat.Name & vbCrLf
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop

    Dim vBodies As Variant
    vBodies = swPart.GetBodies2(0, True)
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

    swModel.ForceRebuild3 False
    swModel.DeleteOLEObjects
    swModel.Save2 True

    Dim msg As String
    msg = "=== Rapport NettoyerFichier ===" & vbCrLf & vbCrLf
    msg = msg & "Features en erreur   : " & nErreurs & vbCrLf
    msg = msg & "Corps solides caches : " & nCorps & vbCrLf & vbCrLf
    If nErreurs > 0 Then
        msg = msg & "Features a corriger :" & vbCrLf & featList & vbCrLf
    End If
    msg = msg & "Modele reconstruit et sauvegarde."

    If nErreurs > 0 Then
        MsgBox msg, vbExclamation, "NettoyerFichier"
    Else
        MsgBox msg, vbInformation, "NettoyerFichier"
    End If

End Sub
