' ==============================================================
' Macro : CreerMiseEnPlan.bas
' Description : Genere une mise en plan avec vues standard
'               (Face, Dessus, Droite, Isometrique)
' Usage : Outils > Macros > Executer (avec piece ouverte)
' ==============================================================

Option Explicit

Dim swApp   As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swDraw  As SldWorks.DrawingDoc

Sub main()

    Set swApp   = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "Aucun document actif.", vbCritical, "CreerMiseEnPlan"
        Exit Sub
    End If
    If swModel.GetType <> 1 And swModel.GetType <> 2 Then
        MsgBox "Ouvrez une piece ou un assemblage d'abord.", vbCritical, "CreerMiseEnPlan"
        Exit Sub
    End If

    Dim sPartPath As String
    sPartPath = swModel.GetPathName

    Dim sFormat  As String
    Dim sEchelle As String
    sFormat  = InputBox("Format de feuille (A4 / A3 / A2 / A1) :", "CreerMiseEnPlan", "A3")
    If sFormat = "" Then Exit Sub
    sEchelle = InputBox("Echelle (ex: 1:1 / 1:2 / 1:5 / 2:1) :", "CreerMiseEnPlan", "1:1")
    If sEchelle = "" Then Exit Sub

    Dim dW As Double
    Dim dH As Double
    Select Case UCase(Trim(sFormat))
        Case "A4": dW = 0.297:  dH = 0.21
        Case "A3": dW = 0.42:   dH = 0.297
        Case "A2": dW = 0.594:  dH = 0.42
        Case "A1": dW = 0.841:  dH = 0.594
        Case Else: dW = 0.42:   dH = 0.297
    End Select

    Dim parts()  As String
    parts = Split(sEchelle, ":")
    Dim dScale As Double
    dScale = 1
    On Error Resume Next
    dScale = CDbl(Trim(parts(0))) / CDbl(Trim(parts(1)))
    If Err.Number <> 0 Or dScale <= 0 Then dScale = 1
    On Error GoTo 0

    Dim sTemplate As String
    sTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplateDrawing)

    Set swDraw = swApp.NewDocument(sTemplate, 12, dW, dH)

    If swDraw Is Nothing Then
        MsgBox "Impossible de creer le document de mise en plan.", vbCritical, "CreerMiseEnPlan"
        Exit Sub
    End If

    Dim xLeft  As Double
    Dim xRight As Double
    Dim yTop   As Double
    Dim yBot   As Double
    xLeft  = dW * 0.25
    xRight = dW * 0.68
    yTop   = dH * 0.7
    yBot   = dH * 0.35

    Dim vFace   As Object
    Dim vDessus As Object
    Dim vDroite As Object
    Dim vIso    As Object

    Set vFace   = swDraw.CreateDrawViewFromModelView3(sPartPath, "*Face", xLeft, yBot, 0)
    Set vDessus = swDraw.CreateDrawViewFromModelView3(sPartPath, "*Dessus", xLeft, yTop, 0)
    Set vDroite = swDraw.CreateDrawViewFromModelView3(sPartPath, "*Droite", xRight, yBot, 0)
    Set vIso    = swDraw.CreateDrawViewFromModelView3(sPartPath, "*Isometrique", xRight, yTop, 0)

    If Not vFace   Is Nothing Then vFace.ScaleDecimal   = dScale
    If Not vDessus Is Nothing Then vDessus.ScaleDecimal = dScale
    If Not vDroite Is Nothing Then vDroite.ScaleDecimal = dScale
    If Not vIso    Is Nothing Then vIso.ScaleDecimal    = dScale * 0.8

    swDraw.ViewZoomtofit2

    Dim sDrawPath As String
    If sPartPath <> "" Then
        Dim dot As Integer
        dot = InStrRev(sPartPath, ".")
        sDrawPath = Left(sPartPath, dot - 1) & ".SLDDRW"
        swDraw.SaveAs sDrawPath
        MsgBox "Mise en plan creee :" & vbCrLf & sDrawPath, vbInformation, "CreerMiseEnPlan"
    Else
        MsgBox "Mise en plan creee. Enregistrez manuellement le plan.", vbInformation, "CreerMiseEnPlan"
    End If

End Sub
