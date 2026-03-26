' ============================================================
' Macro : CreerMiseEnPlan.bas
' Description : Génère automatiquement une mise en plan avec
'               vues standard (Face, Dessus, Droite, Iso)
'               depuis la pièce/assemblage actif
' Usage : Outils > Macros > Exécuter (avec pièce ouverte)
' ============================================================

Option Explicit

Dim swApp   As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swDraw  As SldWorks.DrawingDoc

Sub main()

    Set swApp   = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    ' --- Vérification ---
    If swModel Is Nothing Then
        MsgBox "Aucun document actif.", vbCritical, "CreerMiseEnPlan"
        Exit Sub
    End If
    If swModel.GetType <> 1 And swModel.GetType <> 2 Then
        MsgBox "Ouvrez une pièce ou un assemblage d'abord.", vbCritical, "CreerMiseEnPlan"
        Exit Sub
    End If

    Dim sPartPath As String
    sPartPath = swModel.GetPathName

    ' --- Paramètres utilisateur ---
    Dim sFormat  As String
    Dim sEchelle As String
    sFormat  = InputBox("Format de feuille (A4 / A3 / A2 / A1) :", "CreerMiseEnPlan", "A3")
    If sFormat = "" Then Exit Sub
    sEchelle = InputBox("Échelle (ex: 1:1 / 1:2 / 1:5 / 2:1) :", "CreerMiseEnPlan", "1:1")
    If sEchelle = "" Then Exit Sub

    ' --- Dimensions feuille ---
    Dim dW As Double, dH As Double
    Select Case UCase(Trim(sFormat))
        Case "A4": dW = 0.297:  dH = 0.21
        Case "A3": dW = 0.42:   dH = 0.297
        Case "A2": dW = 0.594:  dH = 0.42
        Case "A1": dW = 0.841:  dH = 0.594
        Case Else: dW = 0.42:   dH = 0.297
    End Select

    ' --- Calcul échelle ---
    Dim parts()  As String
    parts = Split(sEchelle, ":")
    Dim dScale As Double
    On Error Resume Next
    dScale = CDbl(Trim(parts(0))) / CDbl(Trim(parts(1)))
    If Err.Number <> 0 Or dScale <= 0 Then dScale = 1
    On Error GoTo 0

    ' --- Template mise en plan ---
    Dim sTemplate As String
    sTemplate = swApp.GetUserPreferenceStringValue( _
                    swUserPreferenceStringValue_e.swDefaultTemplateDrawing)

    If sTemplate = "" Or Dir(sTemplate) = "" Then
        ' Fallback : template vierge intégré
        sTemplate = swApp.GetExecutablePath & "\lang\french\sheetformat\a3 - landscape.slddrt"
        If Dir(sTemplate) = "" Then sTemplate = ""
    End If

    ' --- Création du document de mise en plan ---
    Dim bErrors  As Boolean
    Set swDraw = swApp.NewDocument(sTemplate, 12, dW, dH)  ' 12 = swDwgPapersUserDefined

    If swDraw Is Nothing Then
        MsgBox "Impossible de créer le document de mise en plan." & vbCrLf & _
               "Vérifiez le template par défaut dans Outils > Options > Emplacements fichiers.", _
               vbCritical, "CreerMiseEnPlan"
        Exit Sub
    End If

    ' --- Calcul positions des vues (en mètres) ---
    '   Disposition : 1ère dièdre européen
    '   [ Dessus  ]  [  Iso  ]
    '   [ Face    ]  [ Droite]
    Dim marge  As Double: marge  = 0.02
    Dim xLeft  As Double: xLeft  = dW * 0.25
    Dim xRight As Double: xRight = dW * 0.68
    Dim yTop   As Double: yTop   = dH * 0.7
    Dim yBot   As Double: yBot   = dH * 0.35

    Dim vFace  As Object
    Dim vDessus As Object
    Dim vDroite As Object
    Dim vIso   As Object

    ' Vue de Face
    Set vFace = swDraw.CreateDrawViewFromModelView3(sPartPath, "*Face", xLeft, yBot, 0)
    If Not vFace Is Nothing Then vFace.ScaleDecimal = dScale

    ' Vue de Dessus (projetée depuis Face)
    Set vDessus = swDraw.CreateDrawViewFromModelView3(sPartPath, "*Dessus", xLeft, yTop, 0)
    If Not vDessus Is Nothing Then vDessus.ScaleDecimal = dScale

    ' Vue de Droite (projetée depuis Face)
    Set vDroite = swDraw.CreateDrawViewFromModelView3(sPartPath, "*Droite", xRight, yBot, 0)
    If Not vDroite Is Nothing Then vDroite.ScaleDecimal = dScale

    ' Vue Isométrique
    Set vIso = swDraw.CreateDrawViewFromModelView3(sPartPath, "*Isométrique", xRight, yTop, 0)
    If Not vIso Is Nothing Then
        vIso.ScaleDecimal = dScale * 0.8   ' légèrement réduite pour lisibilité
    End If

    ' --- Zoom ajusté ---
    swDraw.ViewZoomtofit2

    ' --- Sauvegarde automatique ---
    Dim sDrawPath As String
    If sPartPath <> "" Then
        Dim dot As Integer
        dot = InStrRev(sPartPath, ".")
        sDrawPath = Left(sPartPath, dot - 1) & ".SLDDRW"
        swDraw.SaveAs sDrawPath
        MsgBox "Mise en plan créée :" & vbCrLf & sDrawPath, vbInformation, "CreerMiseEnPlan"
    Else
        MsgBox "Mise en plan créée." & vbCrLf & _
               "La pièce n'est pas sauvegardée, enregistrez manuellement le plan.", _
               vbInformation, "CreerMiseEnPlan"
    End If

End Sub
