Attribute VB_Name = "modLocations"
Option Explicit

Private Function AutoRemisePct(ByVal nbJours As Long) As Double
    Dim seuil As Double, pct As Double
    seuil = NzDbl(ThisWorkbook.Worksheets(SH_CONFIG).Range("B6").Value)
    pct = NzDbl(ThisWorkbook.Worksheets(SH_CONFIG).Range("B5").Value)
    If nbJours >= seuil Then AutoRemisePct = pct Else AutoRemisePct = 0
End Function

Private Function BuildContratNumber(ByVal id As Long) As String
    BuildContratNumber = "CTR-" & Format(Date, "yyyymm") & "-" & Format(id, "0000")
End Function

Public Sub Location_Ajouter()
    Dim wsForm As Worksheet, lo As ListObject, row As ListRow
    Dim id As Long, nbJours As Long
    Dim prixJour As Double, remiseManuelle As Double, remiseAuto As Double, remiseFinale As Double
    Dim brut As Double, remiseMontant As Double, net As Double

    Set wsForm = ThisWorkbook.Worksheets(SH_FORM_LOCATION)
    Set lo = GetTable(SH_LOCATIONS, TB_LOCATIONS)

    If Trim$(wsForm.Range("B4").Value) = "" Or Trim$(wsForm.Range("B5").Value) = "" Then
        MsgBox "ClientID et VehiculeID obligatoires.", vbExclamation
        Exit Sub
    End If

    id = NextId(lo, "LocationID")
    nbJours = DateSafe(wsForm.Range("B7").Value) - DateSafe(wsForm.Range("B6").Value)
    If nbJours <= 0 Then nbJours = 1

    prixJour = NzDbl(wsForm.Range("B8").Value)
    remiseManuelle = NzDbl(wsForm.Range("B9").Value)
    remiseAuto = AutoRemisePct(nbJours)
    remiseFinale = Application.WorksheetFunction.Max(remiseManuelle, remiseAuto)

    brut = nbJours * prixJour
    remiseMontant = brut * remiseFinale / 100
    net = brut - remiseMontant

    Set row = lo.ListRows.Add
    row.Range.Cells(1, lo.ListColumns("LocationID").Index).Value = id
    row.Range.Cells(1, lo.ListColumns("NumeroContrat").Index).Value = BuildContratNumber(id)
    row.Range.Cells(1, lo.ListColumns("ClientID").Index).Value = wsForm.Range("B4").Value
    row.Range.Cells(1, lo.ListColumns("VehiculeID").Index).Value = wsForm.Range("B5").Value
    row.Range.Cells(1, lo.ListColumns("DateDebut").Index).Value = DateSafe(wsForm.Range("B6").Value)
    row.Range.Cells(1, lo.ListColumns("DateFinPrevue").Index).Value = DateSafe(wsForm.Range("B7").Value)
    row.Range.Cells(1, lo.ListColumns("NbJours").Index).Value = nbJours
    row.Range.Cells(1, lo.ListColumns("PrixJourDH").Index).Value = prixJour
    row.Range.Cells(1, lo.ListColumns("RemisePct").Index).Value = remiseFinale
    row.Range.Cells(1, lo.ListColumns("MontantBrut").Index).Value = brut
    row.Range.Cells(1, lo.ListColumns("MontantRemise").Index).Value = remiseMontant
    row.Range.Cells(1, lo.ListColumns("MontantNet").Index).Value = net
    row.Range.Cells(1, lo.ListColumns("TotalPaye").Index).Value = 0
    row.Range.Cells(1, lo.ListColumns("ResteAPayer").Index).Value = net
    row.Range.Cells(1, lo.ListColumns("Statut").Index).Value = "RESERVATION"
    row.Range.Cells(1, lo.ListColumns("EtatDepart").Index).Value = wsForm.Range("B10").Value
    row.Range.Cells(1, lo.ListColumns("EtatRetour").Index).Value = ""
    row.Range.Cells(1, lo.ListColumns("DateCreation").Index).Value = Now

    UpdateVehiculeStatut wsForm.Range("B5").Value, "Réservée"

    If NzDbl(wsForm.Range("B13").Value) > 0 Then
        Paiement_Ajouter id, Date, wsForm.Range("B12").Value, NzDbl(wsForm.Range("B13").Value), wsForm.Range("B14").Value
    End If

    MsgBox "Location créée: " & BuildContratNumber(id), vbInformation
End Sub

Public Sub Location_Modifier()
    Dim wsForm As Worksheet, lo As ListObject, r As ListRow
    Set wsForm = ThisWorkbook.Worksheets(SH_FORM_LOCATION)
    Set lo = GetTable(SH_LOCATIONS, TB_LOCATIONS)

    If Trim$(wsForm.Range("B3").Value) = "" Then
        MsgBox "LocationID obligatoire.", vbExclamation
        Exit Sub
    End If

    Set r = FindRowByValue(lo, "LocationID", wsForm.Range("B3").Value)
    If r Is Nothing Then
        MsgBox "Location introuvable.", vbCritical
        Exit Sub
    End If

    r.Range.Cells(1, lo.ListColumns("DateDebut").Index).Value = DateSafe(wsForm.Range("B6").Value)
    r.Range.Cells(1, lo.ListColumns("DateFinPrevue").Index).Value = DateSafe(wsForm.Range("B7").Value)
    r.Range.Cells(1, lo.ListColumns("EtatDepart").Index).Value = wsForm.Range("B10").Value
    r.Range.Cells(1, lo.ListColumns("EtatRetour").Index).Value = wsForm.Range("B11").Value

    MsgBox "Location modifiée.", vbInformation
End Sub

Public Sub Location_Depart()
    ChangeLocationStatus "DEPART"
End Sub

Public Sub Location_Retour()
    Dim id As Variant
    id = ThisWorkbook.Worksheets(SH_FORM_LOCATION).Range("B3").Value
    ChangeLocationStatus "RETOUR"
    If Trim$(id) <> "" Then
        SetDateRetourEtVehicule CLng(id)
    End If
End Sub

Public Sub Location_Annuler()
    ChangeLocationStatus "ANNULATION"
End Sub

Public Sub Location_Prolonger()
    Dim wsForm As Worksheet, lo As ListObject, r As ListRow
    Dim id As Variant, newFin As Date
    Set wsForm = ThisWorkbook.Worksheets(SH_FORM_LOCATION)
    Set lo = GetTable(SH_LOCATIONS, TB_LOCATIONS)

    id = wsForm.Range("B3").Value
    If Trim$(id) = "" Then
        MsgBox "LocationID obligatoire.", vbExclamation
        Exit Sub
    End If

    Set r = FindRowByValue(lo, "LocationID", id)
    If r Is Nothing Then
        MsgBox "Location introuvable.", vbCritical
        Exit Sub
    End If

    newFin = DateSafe(wsForm.Range("B7").Value)
    If newFin <= r.Range.Cells(1, lo.ListColumns("DateFinPrevue").Index).Value Then
        MsgBox "Nouvelle date fin invalide.", vbExclamation
        Exit Sub
    End If

    r.Range.Cells(1, lo.ListColumns("DateFinPrevue").Index).Value = newFin
    r.Range.Cells(1, lo.ListColumns("Statut").Index).Value = "PROLONGATION"

    MsgBox "Prolongation enregistrée.", vbInformation
End Sub

Private Sub ChangeLocationStatus(ByVal newStatus As String)
    Dim wsForm As Worksheet, lo As ListObject, r As ListRow
    Dim id As Variant, vehiculeId As Variant

    Set wsForm = ThisWorkbook.Worksheets(SH_FORM_LOCATION)
    Set lo = GetTable(SH_LOCATIONS, TB_LOCATIONS)

    id = wsForm.Range("B3").Value
    If Trim$(id) = "" Then
        MsgBox "LocationID obligatoire.", vbExclamation
        Exit Sub
    End If

    Set r = FindRowByValue(lo, "LocationID", id)
    If r Is Nothing Then
        MsgBox "Location introuvable.", vbCritical
        Exit Sub
    End If

    r.Range.Cells(1, lo.ListColumns("Statut").Index).Value = newStatus
    vehiculeId = r.Range.Cells(1, lo.ListColumns("VehiculeID").Index).Value

    Select Case newStatus
        Case "DEPART", "PROLONGATION"
            UpdateVehiculeStatut vehiculeId, "Louée"
        Case "RETOUR", "ANNULATION"
            UpdateVehiculeStatut vehiculeId, "Disponible"
    End Select

    MsgBox "Statut mis à jour: " & newStatus, vbInformation
End Sub

Private Sub SetDateRetourEtVehicule(ByVal locationId As Long)
    Dim loLoc As ListObject, r As ListRow
    Set loLoc = GetTable(SH_LOCATIONS, TB_LOCATIONS)
    Set r = FindRowByValue(loLoc, "LocationID", locationId)
    If r Is Nothing Then Exit Sub

    r.Range.Cells(1, loLoc.ListColumns("DateRetourReelle").Index).Value = Date
    r.Range.Cells(1, loLoc.ListColumns("EtatRetour").Index).Value = ThisWorkbook.Worksheets(SH_FORM_LOCATION).Range("B11").Value
    UpdateVehiculeStatut r.Range.Cells(1, loLoc.ListColumns("VehiculeID").Index).Value, "Disponible"
End Sub

Private Sub UpdateVehiculeStatut(ByVal vehiculeId As Variant, ByVal statut As String)
    Dim lo As ListObject, r As ListRow
    Set lo = GetTable(SH_VEHICULES, TB_VEHICULES)
    Set r = FindRowByValue(lo, "VehiculeID", vehiculeId)
    If Not r Is Nothing Then
        r.Range.Cells(1, lo.ListColumns("Statut").Index).Value = statut
    End If
End Sub
