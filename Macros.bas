Attribute VB_Name = "AssignmentMacros"
Option Explicit

' ---- Macro 1: format the active table (recorded-style) -------------------------------
Sub FormatSalesTable()
    With Selection
        .Font.Name = "Arial"
        .Borders.LineStyle = xlContinuous
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(31, 56, 100)
        .Rows(1).Font.Color = RGB(255, 255, 255)
        .Columns.AutoFit
    End With
End Sub

' ---- Macro 2: highlight sales above a threshold -------------------------------------
Sub HighlightHighSales()
    Dim c As Range, t As Double
    t = Val(InputBox("Highlight values greater than:", "Threshold", 500))
    For Each c In Selection.Cells
        If IsNumeric(c.Value) And Not IsEmpty(c.Value) Then
            If c.Value > t Then c.Interior.Color = RGB(198, 239, 206)
        End If
    Next c
End Sub

' ---- Macro 3: run automatically when the workbook opens (put in ThisWorkbook) -------
Private Sub Workbook_Open()
    Sheets("Index").Activate
    MsgBox "Welcome! Today is " & Format(Date, "dd-mmm-yyyy"), vbInformation, "Excel Assignment"
End Sub

' ---- Macro 4: assign to a button or the shortcut Ctrl+Shift+T -----------------------
Sub TotalSelection()
    MsgBox "Sum of selection = " & Format(Application.WorksheetFunction.Sum(Selection), "#,##0.00")
End Sub
