Attribute VB_Name = "modPaiements"
Option Explicit

Public Sub Paiement_AjouterDepuisForm()
    Dim wsForm As Worksheet
    Set wsForm = ThisWorkbook.Worksheets(SH_FORM_LOCATION)

    If Trim$(wsForm.Range("B3").Value) = "" Then
        MsgBox "LocationID obligatoire.", vbExclamation
        Exit Sub
    End If

    Paiement_Ajouter CLng(wsForm.Range("B3").Value), Date, wsForm.Range("B12").Value, NzDbl(wsForm.Range("B13").Value), wsForm.Range("B14").Value
End Sub

Public Sub Paiement_Ajouter(ByVal locationId As Long, ByVal dt As Date, ByVal mode As String, ByVal montant As Double, ByVal reference As String)
    Dim loP As ListObject, loL As ListObject
    Dim row As ListRow, locRow As ListRow
    Dim totalPaye As Double, net As Double

    If montant <= 0 Then Exit Sub

    Set loP = GetTable(SH_PAIEMENTS, TB_PAIEMENTS)
    Set loL = GetTable(SH_LOCATIONS, TB_LOCATIONS)

    Set row = loP.ListRows.Add
    row.Range.Cells(1, loP.ListColumns("PaiementID").Index).Value = NextId(loP, "PaiementID")
    row.Range.Cells(1, loP.ListColumns("LocationID").Index).Value = locationId
    row.Range.Cells(1, loP.ListColumns("DatePaiement").Index).Value = dt
    row.Range.Cells(1, loP.ListColumns("ModePaiement").Index).Value = mode
    row.Range.Cells(1, loP.ListColumns("MontantDH").Index).Value = montant
    row.Range.Cells(1, loP.ListColumns("Reference").Index).Value = reference

    Set locRow = FindRowByValue(loL, "LocationID", locationId)
    If Not locRow Is Nothing Then
        totalPaye = NzDbl(locRow.Range.Cells(1, loL.ListColumns("TotalPaye").Index).Value) + montant
        net = NzDbl(locRow.Range.Cells(1, loL.ListColumns("MontantNet").Index).Value)
        locRow.Range.Cells(1, loL.ListColumns("TotalPaye").Index).Value = totalPaye
        locRow.Range.Cells(1, loL.ListColumns("ResteAPayer").Index).Value = net - totalPaye
    End If

    MsgBox "Paiement ajouté.", vbInformation
End Sub
