Sub Test()
Application.ScreenUpdating = False

Dim strstart As String
Dim strend As String

'------------Initialisierung------------
    intstart = Cells(3, 10).Value & "*"
    intend = Cells(4, 10).Value & "*"

'------------Sortiere von A bis Z------------
    Columns("A:A").Select
    ActiveWorkbook.Worksheets("Tabelle1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Tabelle1").Sort.SortFields.Add Key:=Range("A1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Tabelle1").Sort
        .SetRange Range("A:A")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'------------Ermittle erstes Element------------
    Set Rng = ActiveSheet.Range("A:A").Find(What:=intstart, LookIn:=xlValues, LookAt:=xlWhole)
    Var = Rng.Row

'------------lösche Zeilen vor erstem Element------------
    ActiveSheet.Range(Cells(1, 1), Cells(Var, 1)).EntireRow.Delete

'------------Sortiere von A bis Z------------
    Columns("A:A").Select
    ActiveWorkbook.Worksheets("Tabelle1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Tabelle1").Sort.SortFields.Add Key:=Range("A1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Tabelle1").Sort
        .SetRange Range("A:A")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'------------Ermittle letztes Element------------
    Set Rng2 = ActiveSheet.Range("A:A").Find(What:=intend, LookIn:=xlValues, LookAt:=xlWhole)
    Var2 = Rng2.Row
    
'------------lösche Zeilen nach letztem Element------------
    ActiveSheet.Range(Cells(1, 1), Cells(Var2, 1)).EntireRow.Delete
    
'------------Sortiere von A bis Z------------
    Columns("A:A").Select
    ActiveWorkbook.Worksheets("Tabelle1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Tabelle1").Sort.SortFields.Add Key:=Range("A1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Tabelle1").Sort
        .SetRange Range("A:A")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'------------Trennen von Nummern und Name------------
    nZeile = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "=LEFT(RC[-1],FIND("" "",RC[-1],2)-1)"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "=RIGHT(RC[-2],LEN(RC[-2])-FIND("" "",RC[-2],2))"
    Range("B1:C1").Select
    Selection.AutoFill Destination:=Range(Cells(1, 2), Cells(nZeile, 3)), Type:=xlFillDefault
    Columns("B:C").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:A").ClearContents
    
'------------Doppelte Einträger löschen------------
    ActiveSheet.Range(Cells(1, 2), Cells(nZeile, 3)).RemoveDuplicates Columns:=1, Header:=xlYes

'------------Überschriften Einfügen------------
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("K1").Value = "CPS"
    Range("L1").Value = "SWC"
    Range("M1").Value = "Funktion"

'------------Schöne Auflistung------------
    Z = 0
    n = 1
    
NeueCps:

        '------------CPS auflisten------------
        Set Rng3 = ActiveSheet.Range("B:B").Find("4." & n, LookIn:=xlValues, LookAt:=xlWhole)
        On Error GoTo Next1
        Cells(Rng3.Row, 3).Select
        Selection.Copy
        Cells(n + Z + 1, 11).Select
        ActiveSheet.Paste
        
        '------------SWCauflisten------------
        For m = 1 To 1000
            Set Rng4 = ActiveSheet.Range("B:B").Find("4." & n & "." & m, LookIn:=xlValues, LookAt:=xlWhole)
            On Error GoTo Next2
            Cells(Rng4.Row, 3).Select
            Selection.Copy
            Cells(n + Z + 1, 12).Select
            ActiveSheet.Paste
            Z = Z + 1
        Next
    
Next2:
n = n + 1
Resume NeueCps
  
Next1:

Application.ScreenUpdating = True
Application.ThisWorkbook.RefreshAll
Columns("K:M").EntireColumn.AutoFit

End Sub
