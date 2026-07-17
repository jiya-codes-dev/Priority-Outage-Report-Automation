' ==========================================================
' Project : Priority Outage Report Automation
' Author  : Jiya Sharma
'
' Description:
' Automates the generation of standardized Priority Outage
' Reports from exported ticket data by importing,
' transforming, cleaning, sorting and formatting records.
'
' NOTE:
' This repository contains a generalized implementation.
' All organization-specific paths, filenames and identifiers
' have been removed.
' ==========================================================
Sub Main()

Application.ScreenUpdating = False

Worksheets("1").Activate
Dim wrb As Workbook, wrb1 As Workbook
Dim wrs As Worksheet, wrs1 As Worksheet
Dim str As String, filter As String
Dim lg As Long, lg2 As Long


Set wrb1 = ThisWorkbook
'Set wrs1 = wrb1.Sheets("MyTest")
Set wrs2 = wrb1.Sheets("1")
wrs2.Range("a2:AB200000").ClearContents
'wrs1.Cells.Clear


Set wrb = Workbooks.Open("C:\MPC\DB\Ctask Priority Report.xlsx")
Set wrs = wrb.Sheets(1)
lg = wrs.Range("A" & wrs.Cells.Rows.Count).End(xlUp).Row

wrs.Range("A2:W" & lg).Copy

wrs2.Range("A2").PasteSpecial Paste:=xlValues
wrs2.Range("A2").PasteSpecial Paste:=xlFormats
Application.CutCopyMode = False

wrb.Close



Worksheets("1").Activate
Columns("F:I").EntireColumn.Delete

Call formula
Call Del
Call nameActionItem
Call Heading
Call Outline
Call Del1
Call TheWall

MsgBox "Your Report is Done Successfully!"
End Sub
Sub formula()

Application.ScreenUpdating = False
Worksheets("1").Activate

lrow2 = Worksheets("1").Cells(Rows.Count, "a").End(xlUp).Row

Worksheets("1").Range("G2:G" & lrow2).Value = "=IF(F2="""","""",LEFT(F2,FIND(CHAR(10),F2,FIND(CHAR(10),F2)+2)))"
Worksheets("1").Range("H2:H" & lrow2).Value = "=IF(F2="""","""",(MID(G2,FIND("")"",G2)+1,256)))"

Worksheets("1").Range("G2:H" & lrow2).Value = Worksheets("1").Range("G2:H" & lrow2).Value
Worksheets("1").Range("I2:I" & lrow2).Value = "=Trim(clean(H2))"

Dim myArea As Range
    For Each myArea In Sheets("1").Rows(1).Areas
        myArea.Value = Evaluate("IF(ROW(" & myArea.Address & "),CLEAN(TRIM(" & myArea.Address & ")))")
    Next myArea

Worksheets("1").Range("I2:I" & lrow2).Value = Worksheets("1").Range("I2:I" & lrow2).Value


End Sub
Sub CleanTrim()
    
End Sub
Sub Del()
Worksheets("1").Activate
Columns("F:H").EntireColumn.Delete
End Sub
Sub nameActionItem()
Worksheets("1").Activate
Worksheets("1").Range("F1").Value = "Action Item"
Worksheets("1").Range("F1").Font.Bold = True

End Sub
Sub Heading()
'
' Macro2 Macro
'
' Keyboard Shortcut: Ctrl+m
'
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "CTASK"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Change"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Short Description"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Type"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "State"
    Range("A1:F1").Select
    
    With Selection.Font
        .Color = -6736897
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("A1:F1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16771071
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
End Sub

Sub Outline()
'
' Outline Macro
'
' Keyboard Shortcut: Ctrl+o
'
    
    Range("A1:F10").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
End Sub
Sub Del1()
Worksheets("1").Activate
Columns("G:P").EntireColumn.Delete
End Sub

Sub TheWall()
    With ActiveSheet.UsedRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End Sub
