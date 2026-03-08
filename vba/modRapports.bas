Attribute VB_Name = "modRapports"
Option Explicit

Public Sub Rapport_Impayes()
    Dim ws As Worksheet, lo As ListObject, r As ListRow, outRow As Long
    Set ws = GetOrCreateSheet("RAPPORT_IMPAYES")
    Set lo = GetTable(SH_LOCATIONS, TB_LOCATIONS)

    ws.Cells.Clear
    ws.Range("A1:H1").Value = Array("Contrat", "ClientID", "VehiculeID", "DateDebut", "DateFin", "MontantNet", "Payé", "Reste")
    outRow = 2

    For Each r In lo.ListRows
        If NzDbl(r.Range.Cells(1, lo.ListColumns("ResteAPayer").Index).Value) > 0 Then
            ws.Cells(outRow, "A").Value = r.Range.Cells(1, lo.ListColumns("NumeroContrat").Index).Value
            ws.Cells(outRow, "B").Value = r.Range.Cells(1, lo.ListColumns("ClientID").Index).Value
            ws.Cells(outRow, "C").Value = r.Range.Cells(1, lo.ListColumns("VehiculeID").Index).Value
            ws.Cells(outRow, "D").Value = r.Range.Cells(1, lo.ListColumns("DateDebut").Index).Value
            ws.Cells(outRow, "E").Value = r.Range.Cells(1, lo.ListColumns("DateFinPrevue").Index).Value
            ws.Cells(outRow, "F").Value = r.Range.Cells(1, lo.ListColumns("MontantNet").Index).Value
            ws.Cells(outRow, "G").Value = r.Range.Cells(1, lo.ListColumns("TotalPaye").Index).Value
            ws.Cells(outRow, "H").Value = r.Range.Cells(1, lo.ListColumns("ResteAPayer").Index).Value
            outRow = outRow + 1
        End If
    Next r

    ws.Columns.AutoFit
    MsgBox "Rapport impayés généré.", vbInformation
End Sub

Public Sub Rapport_Retards()
    Dim ws As Worksheet, lo As ListObject, r As ListRow, outRow As Long, st As String
    Set ws = GetOrCreateSheet("RAPPORT_RETARDS")
    Set lo = GetTable(SH_LOCATIONS, TB_LOCATIONS)

    ws.Cells.Clear
    ws.Range("A1:G1").Value = Array("Contrat", "ClientID", "VehiculeID", "DateFinPrévue", "DateRetourRéelle", "Statut", "JoursRetard")
    outRow = 2

    For Each r In lo.ListRows
        st = UCase$(CStr(r.Range.Cells(1, lo.ListColumns("Statut").Index).Value))
        If st = "DEPART" Or st = "PROLONGATION" Then
            If Date > DateSafe(r.Range.Cells(1, lo.ListColumns("DateFinPrevue").Index).Value) Then
                ws.Cells(outRow, "A").Value = r.Range.Cells(1, lo.ListColumns("NumeroContrat").Index).Value
                ws.Cells(outRow, "B").Value = r.Range.Cells(1, lo.ListColumns("ClientID").Index).Value
                ws.Cells(outRow, "C").Value = r.Range.Cells(1, lo.ListColumns("VehiculeID").Index).Value
                ws.Cells(outRow, "D").Value = r.Range.Cells(1, lo.ListColumns("DateFinPrevue").Index).Value
                ws.Cells(outRow, "E").Value = r.Range.Cells(1, lo.ListColumns("DateRetourReelle").Index).Value
                ws.Cells(outRow, "F").Value = st
                ws.Cells(outRow, "G").Value = Date - DateSafe(r.Range.Cells(1, lo.ListColumns("DateFinPrevue").Index).Value)
                outRow = outRow + 1
            End If
        End If
    Next r

    ws.Columns.AutoFit
    MsgBox "Rapport retards généré.", vbInformation
End Sub
