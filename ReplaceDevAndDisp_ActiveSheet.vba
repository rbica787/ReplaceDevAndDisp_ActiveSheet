Sub ReplaceDevAndDisp_ActiveSheet()
    Dim ws As Worksheet
    Dim rng As Range

    Set ws = ActiveSheet

    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo Cleanup

    Set rng = ws.UsedRange
    rng.Replace What:="_dev", Replacement:=".dvtx", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False
    rng.Replace What:="_disp", Replacement:=".dspx", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
End Sub
