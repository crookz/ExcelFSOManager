Attribute VB_Name = "Feature_List_Tools"
Sub Jump_to_Feature()
'
' Jump_to_Feature Macro
'

    Dim Feature As String
    Dim FeatureLookup As Boolean: FeatureLookup = False
    Dim CurrentRow As Range
    
    'Refresh pivot table and deal silently with non-critical errors
    On Error Resume Next
    ActiveSheet.PivotTables(1).PivotCache.Refresh
    On Error GoTo 0
    If Err.Number <> 1004 And Err.Number <> 0 Then
        MsgBox ("CONTACT OWNER: Unknown Pivot Table Error Has Occured, CODE: " & Err.Number)
    End If
    
    'Check only one cell is selected and its not empty
    If Selection.Rows.Count + Selection.Columns.Count = 2 Then
        Feature = Selection
        Sheets("Game Features").Select
        If Feature <> "" Then
        
            'Cycle through each feature until you find a match else error
            For Each CurrentRow In Range("Table_GameFeatures[Features]")
                If Feature = CurrentRow.Cells(1, 1).Value Then
                    CurrentRow.Select
                    FeatureLookup = True
                    Exit For
                End If
            Next CurrentRow
            
            If Not FeatureLookup Then
                MsgBox ("Sorry, The feature is not present. Try refreshing the pivot table.")
            End If
            
        Else
            MsgBox ("ERROR: You did not select a Feature.")
        End If
        
    Else
        MsgBox ("ERROR: You selected more than One Feature! Please select one cell only.")
    End If
    
End Sub

Sub Reset_Summary_Pivot()
    On Error Resume Next
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
    On Error GoTo 0
    
    If Err.Number <> 1004 And Err.Number <> 0 Then
        MsgBox ("CONTACT OWNER: Unknown Pivot Table Error Has Occured, CODE: " & Err.Number)
    End If
End Sub
    
Sub Add_Feature_Header()
'
'Adds a grey header from the category of the selected feature.
'
    Dim Category As String
    Dim InsertRow As Long
    'If a single row is selected (or cell)
    If Selection.Count = 1 Or Selection.Rows.Count = 1 Then
        'Select & copy the category cell on this row
        InsertRow = ActiveCell.Row
        Category = Range(ConvertToLetter(Range("Table_GameFeatures[Category]").Column) & InsertRow).Value
        'Insert new header row (1 in status colum) with Title as Category
        Cells(InsertRow, 1).EntireRow.Insert
        Range(ConvertToLetter(Range("Table_GameFeatures[Features]").Column) & InsertRow).FormulaR1C1 = Category
        Range(ConvertToLetter(Range("Table_GameFeatures[STATUS]").Column) & InsertRow).FormulaR1C1 = "1"
    Else
        MsgBox "You must have ONE CELL/ROW selected."
    End If

End Sub

Function Inbox_Reader() As Range
'
' Cycles through Inbox until Name is Blank. First line with an empty status is read.
'

    Dim targetSheet As Worksheet: Set targetSheet = Worksheets("Inbox")
    Dim CurrentRow As Range
    Dim RecordData As Range
    Dim RowCount As Integer
    
    RowCount = 0
    
    For Each CurrentRow In targetSheet.Rows
    
        ' Exit on first empty Feature name (Assumed end of sheet)
         If targetSheet.Cells(CurrentRow.Row, targetSheet.Range("InboxFeatures[Name]").Column).Value = "" Then
            MsgBox ("SORRY: The Inbox Sheet currently has no features to process.")
        Set Inbox_Reader = Nothing
            Exit For
        End If
      
        ' When feature isnt empty, if Status is empty, is assumed as not processed
        If targetSheet.Cells(CurrentRow.Row, targetSheet.Range("InboxFeatures[Status]").Column).Value = "" Then
            'Grab Row set status to YES and exit
            targetSheet.Cells(CurrentRow.Row, targetSheet.Range("InboxFeatures[Status]").Column).FormulaR1C1 = "YES"
            Set RecordData = CurrentRow
            MsgBox ("Feature: " & RecordData.Cells(1, targetSheet.Range("InboxFeatures[Name]").Column).Value & " has been processed. STUDIO: " & RecordData.Cells(1, targetSheet.Range("InboxFeatures[Studio]").Column).Value)
            'Ans = MsgBox(Msg, vbQuestion + vbYesNoCancel)
            Set Inbox_Reader = RecordData
            Exit For
        End If
    
        RowCount = RowCount + 1
        
    Next CurrentRow
End Function

Sub Feature_Insert()
'
' Feature_Insert Macro
' Insert an Inbo Feature at current location
'
' Keyboard Shortcut: Ctrl+Shift+I
'
    Dim InsertRow As Integer
    Dim InboxFeature As Range
    
    If Selection.Count = 1 Or Selection.Rows.Count = 1 Then
        
        Set InboxFeature = Inbox_Reader()
        If Not InboxFeature Is Nothing Then
        
            'add the new row
            ActiveCell.EntireRow.Insert
            InsertRow = ActiveCell.Row
            
            'Insert category from below cell
            Range(ConvertToLetter(Range("Table_GameFeatures[Category]").Column) & (ActiveCell.Row)).Select
            ActiveCell.FormulaR1C1 = Selection.Offset(1, 0)
            
            'Set Component to gameplay as default
            Range(ConvertToLetter(Range("Table_GameFeatures[Component]").Column) & (ActiveCell.Row)).Select
            ActiveCell.FormulaR1C1 = "Gameplay"
            
            'Set Feature status to APPROVED as default
            Range(ConvertToLetter(Range("Table_GameFeatures[Feature status]").Column) & (ActiveCell.Row)).Select
            ActiveCell.FormulaR1C1 = "APPROVED"
            
            'Set Feature Type to CORE as default
            Range(ConvertToLetter(Range("Table_GameFeatures[Feature Type]").Column) & (ActiveCell.Row)).Select
            ActiveCell.FormulaR1C1 = "CORE"
                 
            'Set studio to MTL as default
            Range(ConvertToLetter(Range("Table_GameFeatures[MTL]").Column) & (ActiveCell.Row)).Select
            ActiveCell.FormulaR1C1 = "2"
            
            'Set all platforms as yes default
            Range(ConvertToLetter(Range("Table_GameFeatures[xbox_one]").Column) & (ActiveCell.Row)).Select
            ActiveCell.FormulaR1C1 = "2"
            ActiveCell.Offset(0, 1).Activate
            ActiveCell.FormulaR1C1 = "2"
            ActiveCell.Offset(0, 1).Activate
            ActiveCell.FormulaR1C1 = "2"
            
            'Set Feature Value
            Range(ConvertToLetter(Range("Table_GameFeatures[Features]").Column) & (ActiveCell.Row)).Select
            ActiveCell.FormulaR1C1 = InboxFeature.Cells(1, 2).Value
            
            'Insert FSO from below cell if recieved value is blank
'            Range(ConvertToLetter(Range("Table_GameFeatures[fso_doc]").Column) & (ActiveCell.Row)).Select
'            If InboxFeature.Cells(1, 3).Value = "" Then
'                ActiveCell.FormulaR1C1 = Selection.Offset(1, 0)
'            Else
'                ActiveCell.FormulaR1C1 = InboxFeature.Cells(1, 3).Value
'            End If
            
            'Set Feature Description
            Range(ConvertToLetter(Range("Table_GameFeatures[Definition]").Column) & (ActiveCell.Row)).Select
            ActiveCell.FormulaR1C1 = InboxFeature.Cells(1, 4).Value
        End If
    Else
        MsgBox "You must only have ONE ROW selected."
    End If
End Sub

Sub GroupHeaders()
'
' Regroups All Headers correctly
'
    'Clean existing groups and set pointer to first category
    Range("Table_GameFeatures").Select
    On Error Resume Next
    Selection.Rows.Ungroup
    On Error GoTo 0
    Range(ConvertToLetter(Range("Table_GameFeatures[Category]").Column) & "4").Select
    
    Dim Start As String: Start = ActiveCell.Row
    Dim i As Long: i = Start
    Dim j As Long: j = 0
    Dim Value As String: Value = ActiveCell.Value
    
    'loop through for each header
    Do While j < Application.WorksheetFunction.CountIf(Range("Table_GameFeatures[STATUS]"), "1")
        'loop through for each matching category
        Do While Range("H" & i).Value = Value
            i = i + 1
        Loop
        'select range and group
        Range("H" & Start & ":H" & i - 1).EntireRow.Select
        Selection.Group
        'skip the header line
        Range("H" & i - 1).Select
        ActiveCell.Offset(2, 0).Activate
        'reset values for next grouping
        i = ActiveCell.Row
        Start = ActiveCell.Row
        Value = ActiveCell.Value
        j = j + 1
    Loop
    
End Sub

Sub FiterByFSO()
'
' Filters featurelist based on selected FSO name
'
    Dim rngSourceColumn         As Range:   Set rngSourceColumn = ActiveSheet.Range("Table_FSOList[Filename]")
    Dim rngTargetColumn         As Range:   Set rngTargetColumn = ActiveWorkbook.Sheets("Game Features").Range("Table_GameFeatures[fso_doc]")
    Dim strFSOname              As String:  strFSOname = rngSourceColumn.Cells(ActiveCell.Row - 1)
    
    ActiveWorkbook.Sheets("Game Features").ListObjects("Table_GameFeatures").Range.AutoFilter Field:=rngTargetColumn.Column, _
        Criteria1:="=*" & strFSOname & "*", Operator:=xlAnd
    
    ActiveWorkbook.Sheets("Game Features").Activate
    
End Sub

