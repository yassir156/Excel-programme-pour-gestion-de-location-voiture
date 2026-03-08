Attribute VB_Name = "modDashboard"
Option Explicit

Public Sub Dashboard_Refresh()
    Dim ws As Worksheet, lo As ListObject, r As ListRow
    Dim ca As Double, paye As Double, reste As Double
    Dim active As Long, reservations As Long, retards As Long

    Set ws = ThisWorkbook.Worksheets(SH_DASHBOARD)
    Set lo = GetTable(SH_LOCATIONS, TB_LOCATIONS)

    ca = 0: paye = 0: reste = 0
    active = 0: reservations = 0: retards = 0

    For Each r In lo.ListRows
        ca = ca + NzDbl(r.Range.Cells(1, lo.ListColumns("MontantNet").Index).Value)
        paye = paye + NzDbl(r.Range.Cells(1, lo.ListColumns("TotalPaye").Index).Value)
        reste = reste + NzDbl(r.Range.Cells(1, lo.ListColumns("ResteAPayer").Index).Value)

        Select Case UCase$(CStr(r.Range.Cells(1, lo.ListColumns("Statut").Index).Value))
            Case "DEPART", "PROLONGATION"
                active = active + 1
                If Date > DateSafe(r.Range.Cells(1, lo.ListColumns("DateFinPrevue").Index).Value) Then
                    retards = retards + 1
                End If
            Case "RESERVATION"
                reservations = reservations + 1
        End Select
    Next r

    ws.Range("B3").Value = ca
    ws.Range("B4").Value = paye
    ws.Range("B5").Value = reste
    ws.Range("B6").Value = active
    ws.Range("B7").Value = reservations
    ws.Range("B8").Value = retards

    Dashboard_FillTopVehicules
    Dashboard_RefreshAlertesEntretien

    MsgBox "Dashboard mis à jour.", vbInformation
End Sub

Private Sub Dashboard_FillTopVehicules()
    Dim ws As Worksheet, loV As ListObject, loL As ListObject
    Dim rV As ListRow, rL As ListRow, outRow As Long
    Dim id As Variant, immat As String, nb As Long, ca As Double

    Set ws = ThisWorkbook.Worksheets(SH_DASHBOARD)
    Set loV = GetTable(SH_VEHICULES, TB_VEHICULES)
    Set loL = GetTable(SH_LOCATIONS, TB_LOCATIONS)

    ws.Range("A12:D200").ClearContents
    outRow = 12

    For Each rV In loV.ListRows
        id = rV.Range.Cells(1, loV.ListColumns("VehiculeID").Index).Value
        immat = CStr(rV.Range.Cells(1, loV.ListColumns("Immatriculation").Index).Value)
        nb = 0: ca = 0

        For Each rL In loL.ListRows
            If CStr(rL.Range.Cells(1, loL.ListColumns("VehiculeID").Index).Value) = CStr(id) Then
                nb = nb + 1
                ca = ca + NzDbl(rL.Range.Cells(1, loL.ListColumns("MontantNet").Index).Value)
            End If
        Next rL

        If nb > 0 Then
            ws.Cells(outRow, "A").Value = id
            ws.Cells(outRow, "B").Value = immat
            ws.Cells(outRow, "C").Value = nb
            ws.Cells(outRow, "D").Value = ca
            outRow = outRow + 1
        End If
    Next rV
End Sub

Private Sub Dashboard_RefreshAlertesEntretien()
    Dim lo As ListObject, r As ListRow
    Set lo = GetTable(SH_ENTRETIEN, TB_ENTRETIEN)

    For Each r In lo.ListRows
        If IsDate(r.Range.Cells(1, lo.ListColumns("DateProchaine").Index).Value) Then
            If Date >= CDate(r.Range.Cells(1, lo.ListColumns("DateProchaine").Index).Value) Then
                r.Range.Cells(1, lo.ListColumns("Alerte").Index).Value = "ROUGE"
            Else
                r.Range.Cells(1, lo.ListColumns("Alerte").Index).Value = "OK"
            End If
        Else
            r.Range.Cells(1, lo.ListColumns("Alerte").Index).Value = "OK"
        End If
    Next r
End Sub
