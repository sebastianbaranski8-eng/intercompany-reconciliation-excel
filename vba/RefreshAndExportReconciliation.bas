Attribute VB_Name = "Module1"
Option Explicit

Sub RefreshAndExportReconciliation()
    Dim cn As WorkbookConnection
    Dim sh As Worksheet
    Dim pt As PivotTable
    Dim ws As Worksheet
    Dim folder As String
    Dim pdfPath As String

    On Error GoTo ErrHandler
    Application.StatusBar = "Refreshing data..."

    ' 1. Refresh queries in the foreground so the macro waits for them
    For Each cn In ThisWorkbook.Connections
        If cn.Type = xlConnectionTypeOLEDB Then cn.OLEDBConnection.BackgroundQuery = False
    Next cn
    ThisWorkbook.RefreshAll

    ' 2. Refresh every pivot table and recalculate
    For Each sh In ThisWorkbook.Worksheets
        For Each pt In sh.PivotTables
            pt.RefreshTable
        Next pt
    Next sh
    Application.Calculate

    ' 3. Check that the reconciliation closes
    Set ws = ThisWorkbook.Worksheets("Reconciliation")
    If ws.Range("C15").Value <> "OK" Then
        If MsgBox("Reconciliation check is not OK. Export anyway?", vbYesNo + vbExclamation) = vbNo Then GoTo CleanExit
    End If

    ' Fit row heights to wrapped text
    ws.UsedRange.Rows.AutoFit

    ' 4. Page setup: landscape, one page
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

    ' 5. Export to PDF next to the workbook (or to Documents for OneDrive files)
    folder = ThisWorkbook.Path
    If folder = "" Or LCase(Left(folder, 4)) = "http" Then folder = Environ("USERPROFILE") & "\Documents"
    pdfPath = folder & "\IC_Reconciliation_" & Format(Date, "yyyy-mm-dd") & ".pdf"

    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, _
        Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=True

    MsgBox "Reconciliation exported to:" & vbCrLf & pdfPath, vbInformation

CleanExit:
    Application.StatusBar = False
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub
