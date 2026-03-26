' ==============================================================
' Macro : SetupPiece.bas
' Description : Remplit les proprietes standard d'une piece SW
' Usage : Outils > Macros > Executer (avec piece ouverte)
' ==============================================================

Option Explicit

Dim swApp      As SldWorks.SldWorks
Dim swModel    As SldWorks.ModelDoc2
Dim swCustProp As SldWorks.CustomPropertyManager

Sub main()

    Set swApp   = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "Aucun document actif. Ouvrez une piece ou un assemblage.", vbCritical, "SetupPiece"
        Exit Sub
    End If

    Dim docType As Integer
    docType = swModel.GetType
    If docType <> 1 And docType <> 2 Then
        MsgBox "Ce document n'est pas une piece ou un assemblage.", vbCritical, "SetupPiece"
        Exit Sub
    End If

    Dim sRef      As String
    Dim sDesig    As String
    Dim sMatiere  As String
    Dim sAuteur   As String
    Dim sProjet   As String
    Dim sRevision As String

    sRef      = InputBox("Reference piece (ex: PRJ-001-001) :", "SetupPiece", "")
    If sRef = "" Then Exit Sub

    sDesig    = InputBox("Designation :", "SetupPiece", "")
    sMatiere  = InputBox("Matiere (ex: Acier S235, Alu 6061, POM...) :", "SetupPiece", "")
    sAuteur   = InputBox("Auteur / Dessinateur :", "SetupPiece", Environ("USERNAME"))
    sProjet   = InputBox("Nom du projet :", "SetupPiece", "")
    sRevision = InputBox("Indice de revision :", "SetupPiece", "A")

    Set swCustProp = swModel.Extension.CustomPropertyManager("")

    Call SetProp(swCustProp, "Reference",   swCustomInfoText, sRef)
    Call SetProp(swCustProp, "Designation", swCustomInfoText, sDesig)
    Call SetProp(swCustProp, "Matiere",     swCustomInfoText, sMatiere)
    Call SetProp(swCustProp, "Auteur",      swCustomInfoText, sAuteur)
    Call SetProp(swCustProp, "Projet",      swCustomInfoText, sProjet)
    Call SetProp(swCustProp, "Revision",    swCustomInfoText, sRevision)
    Call SetProp(swCustProp, "Date",        swCustomInfoText, Format(Now, "dd/mm/yyyy"))
    Call SetProp(swCustProp, "Statut",      swCustomInfoText, "En cours")

    swModel.Save2 True

    Dim msg As String
    msg = "Proprietes enregistrees !" & vbCrLf & vbCrLf
    msg = msg & "Reference  : " & sRef & vbCrLf
    msg = msg & "Designation: " & sDesig & vbCrLf
    msg = msg & "Matiere    : " & sMatiere & vbCrLf
    msg = msg & "Auteur     : " & sAuteur & vbCrLf
    msg = msg & "Revision   : " & sRevision
    MsgBox msg, vbInformation, "SetupPiece"

End Sub

Private Sub SetProp(mgr As SldWorks.CustomPropertyManager, sName As String, iType As Integer, sVal As String)
    mgr.Delete sName
    mgr.Add2   sName, iType, sVal
End Sub
