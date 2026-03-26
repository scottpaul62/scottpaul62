' ============================================================
' Macro : SetupPiece.bas
' Description : Remplit automatiquement les propriétés standard
'               d'une pièce ou assemblage SolidWorks
' Usage : Outils > Macros > Exécuter (avec pièce ouverte)
' ============================================================

Option Explicit

Dim swApp       As SldWorks.SldWorks
Dim swModel     As SldWorks.ModelDoc2
Dim swCustProp  As SldWorks.CustomPropertyManager

Sub main()

    Set swApp   = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    ' --- Vérification document actif ---
    If swModel Is Nothing Then
        MsgBox "Aucun document actif." & vbCrLf & \
               "Ouvrez une pièce ou un assemblage.", vbCritical, "SetupPiece"
        Exit Sub
    End If

    Dim docType As Integer
    docType = swModel.GetType
    If docType <> 1 And docType <> 2 Then   ' 1=Part  2=Assembly
        MsgBox "Ce document n'est pas une pièce ou un assemblage.", vbCritical, "SetupPiece"
        Exit Sub
    End If

    ' --- Saisie des informations ---
    Dim sRef     As String
    Dim sDesig   As String
    Dim sMatiere As String
    Dim sAuteur  As String
    Dim sProjet  As String
    Dim sRevision As String

    sRef      = InputBox("Référence pièce (ex: PRJ-001-001) :", "SetupPiece", "")
    If sRef = "" Then Exit Sub

    sDesig    = InputBox("Désignation :", "SetupPiece", "")
    sMatiere  = InputBox("Matière (ex: Acier S235, Alu 6061, POM...) :", "SetupPiece", "")
    sAuteur   = InputBox("Auteur / Dessinateur :", "SetupPiece", Environ("USERNAME"))
    sProjet   = InputBox("Nom du projet :", "SetupPiece", "")
    sRevision = InputBox("Indice de révision :", "SetupPiece", "A")

    ' --- Écriture des propriétés personnalisées ---
    Set swCustProp = swModel.Extension.CustomPropertyManager("")

    Call SetProp(swCustProp, "Référence",   swCustomInfoText, sRef)
    Call SetProp(swCustProp, "Désignation", swCustomInfoText, sDesig)
    Call SetProp(swCustProp, "Matière",     swCustomInfoText, sMatiere)
    Call SetProp(swCustProp, "Auteur",      swCustomInfoText, sAuteur)
    Call SetProp(swCustProp, "Projet",      swCustomInfoText, sProjet)
    Call SetProp(swCustProp, "Révision",    swCustomInfoText, sRevision)
    Call SetProp(swCustProp, "Date",        swCustomInfoText, Format(Now, "dd/mm/yyyy"))
    Call SetProp(swCustProp, "Statut",      swCustomInfoText, "En cours")

    ' --- Sauvegarde ---
    swModel.Save2 True

    MsgBox "Propriétés enregistrées !" & vbCrLf & vbCrLf & _
           "Référence  : " & sRef     & vbCrLf & _
           "Désignation: " & sDesig   & vbCrLf & _
           "Matière    : " & sMatiere & vbCrLf & _
           "Auteur     : " & sAuteur  & vbCrLf & _
           "Révision   : " & sRevision, vbInformation, "SetupPiece"

End Sub

' Ajoute ou met à jour une propriété custom
Private Sub SetProp(mgr As SldWorks.CustomPropertyManager, _
                    sName As String, iType As Integer, sVal As String)
    mgr.Delete sName
    mgr.Add2   sName, iType, sVal
End Sub
