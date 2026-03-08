Attribute VB_Name = "modVehicules"
Option Explicit

Public Sub Vehicule_Ajouter()
    Dim wsForm As Worksheet, lo As ListObject, row As ListRow
    Set wsForm = ThisWorkbook.Worksheets(SH_FORM_VEHICULE)
    Set lo = GetTable(SH_VEHICULES, TB_VEHICULES)

    If Trim$(wsForm.Range("B4").Value) = "" Then
        MsgBox "Immatriculation obligatoire.", vbExclamation
        Exit Sub
    End If

    Set row = lo.ListRows.Add
    row.Range.Cells(1, lo.ListColumns("VehiculeID").Index).Value = NextId(lo, "VehiculeID")
    row.Range.Cells(1, lo.ListColumns("Immatriculation").Index).Value = wsForm.Range("B4").Value
    row.Range.Cells(1, lo.ListColumns("Marque").Index).Value = wsForm.Range("B5").Value
    row.Range.Cells(1, lo.ListColumns("Modele").Index).Value = wsForm.Range("B6").Value
    row.Range.Cells(1, lo.ListColumns("Annee").Index).Value = wsForm.Range("B7").Value
    row.Range.Cells(1, lo.ListColumns("Km").Index).Value = wsForm.Range("B8").Value
    row.Range.Cells(1, lo.ListColumns("Carburant").Index).Value = wsForm.Range("B9").Value
    row.Range.Cells(1, lo.ListColumns("PrixJourDH").Index).Value = NzDbl(wsForm.Range("B10").Value)
    row.Range.Cells(1, lo.ListColumns("Statut").Index).Value = IIf(Trim$(wsForm.Range("B11").Value) = "", "Disponible", wsForm.Range("B11").Value)
    row.Range.Cells(1, lo.ListColumns("DateAjout").Index).Value = Date

    MsgBox "Véhicule ajouté.", vbInformation
End Sub

Public Sub Vehicule_Modifier()
    Dim wsForm As Worksheet, lo As ListObject, r As ListRow
    Set wsForm = ThisWorkbook.Worksheets(SH_FORM_VEHICULE)
    Set lo = GetTable(SH_VEHICULES, TB_VEHICULES)

    If Trim$(wsForm.Range("B3").Value) = "" Then
        MsgBox "VehiculeID obligatoire pour modifier.", vbExclamation
        Exit Sub
    End If

    Set r = FindRowByValue(lo, "VehiculeID", wsForm.Range("B3").Value)
    If r Is Nothing Then
        MsgBox "Véhicule introuvable.", vbCritical
        Exit Sub
    End If

    r.Range.Cells(1, lo.ListColumns("Immatriculation").Index).Value = wsForm.Range("B4").Value
    r.Range.Cells(1, lo.ListColumns("Marque").Index).Value = wsForm.Range("B5").Value
    r.Range.Cells(1, lo.ListColumns("Modele").Index).Value = wsForm.Range("B6").Value
    r.Range.Cells(1, lo.ListColumns("Annee").Index).Value = wsForm.Range("B7").Value
    r.Range.Cells(1, lo.ListColumns("Km").Index).Value = wsForm.Range("B8").Value
    r.Range.Cells(1, lo.ListColumns("Carburant").Index).Value = wsForm.Range("B9").Value
    r.Range.Cells(1, lo.ListColumns("PrixJourDH").Index).Value = NzDbl(wsForm.Range("B10").Value)
    r.Range.Cells(1, lo.ListColumns("Statut").Index).Value = wsForm.Range("B11").Value

    MsgBox "Véhicule modifié.", vbInformation
End Sub

Public Sub Vehicule_Supprimer()
    Dim wsForm As Worksheet, lo As ListObject, r As ListRow
    Set wsForm = ThisWorkbook.Worksheets(SH_FORM_VEHICULE)
    Set lo = GetTable(SH_VEHICULES, TB_VEHICULES)

    If Trim$(wsForm.Range("B3").Value) = "" Then
        MsgBox "VehiculeID obligatoire pour supprimer.", vbExclamation
        Exit Sub
    End If

    Set r = FindRowByValue(lo, "VehiculeID", wsForm.Range("B3").Value)
    If r Is Nothing Then
        MsgBox "Véhicule introuvable.", vbCritical
        Exit Sub
    End If

    If MsgBox("Supprimer ce véhicule ?", vbQuestion + vbYesNo) = vbYes Then
        r.Delete
        MsgBox "Véhicule supprimé.", vbInformation
    End If
End Sub
