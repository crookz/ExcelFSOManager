Attribute VB_Name = "Prism_Tools"
'**********************VBA LOGGING SYSTEM*******************************************
'Author: Sean Crooks
'Series of tools for interfacing with prims Reporting System
'
'***********************************************************************************
Enum prismFeatureData
    Features = 1
    Catagory = 2
    Complete = 3
    Level = 4
    Status = 5
    Days = 6
    Studio = 7
End Enum

Enum prismFsoData
    Fsos = 1
    Progress = 2
    JIRA = 3
    Studio = 4
    Status = 5
    Estimate = 6
    Replication = 7
    Lines = 8
End Enum

Enum prismFeatureColumns
    Features = 2
    Catagory = 1
    Studio = 3
    L0_Complete = 6
    L0_Sprint = 7
    L0_Level = 101
    L0_Status = 4
    L0_Days = 5
    L1_Complete = 10
    L1_Sprint = 11
    L1_Level = 100
    L1_Status = 8
    L1_Days = 9
End Enum

Enum prismFsoColumns
    EstimateType = 1
    FsoName = 2
    Studio = 3
    L1_Status = 8
    L1_Days = 9
    L1_Progress = 10
    L1_Sprint = 11
    JIRA = 22
    Replication = 23
End Enum


Option Explicit

Function CollectFsoData(targetSheet As Worksheet) As Variant
    Dim FsoArray() As Variant
    Dim ProgressArray() As Variant
    Dim JIRAArray() As Variant
    Dim StudioArray() As Variant
    Dim StatusArray() As Variant
    Dim EstimateTypeArray() As Variant
    Dim ReplicationArray() As Variant
    Dim TotalSPLinesArray() As Variant
    Dim TotalMPLinesArray() As Variant
        
    Dim fso_loop As Integer: fso_loop = 1
    Dim fso_count As Integer: fso_count = 1
    Dim CollectedData() As Variant

    'grab required date to arrays d

    FsoArray = targetSheet.Range("Table_FSOList[Summary]").Value
    ProgressArray = targetSheet.Range("Table_FSOList[Percentage Combination]").Value
    JIRAArray = targetSheet.Range("Table_FSOList[Key]").Value
    StudioArray = targetSheet.Range("Table_FSOList[FSO Studio Owner]").Value
    StatusArray = targetSheet.Range("Table_FSOList[Status]").Value
    EstimateTypeArray = targetSheet.Range("Table_FSOList[Estimate Type]").Value
    ReplicationArray = targetSheet.Range("Table_FSOList[Replication]").Value
    'Added number of lines to allow weighting per FSO
    TotalSPLinesArray = targetSheet.Range("Table_FSOList[SP FSO Total]").Value
    TotalMPLinesArray = targetSheet.Range("Table_FSOList[MP FSO Total]").Value
    
    ReDim CollectedData(1 To 8, LBound(FsoArray) To UBound(FsoArray)) 'TODO autocount enum total
    
    'cleans data for prism entry by:
    ' - Removing Cut features
    ' - Removing dataless features (empty progression)
    
    For fso_loop = LBound(FsoArray) To UBound(FsoArray)
    
        If FsoArray(fso_loop, 1) <> "" And StatusArray(fso_loop, 1) <> "FSO - Cut" And ProgressArray(fso_loop, 1) <> "" Then
            
            CollectedData(prismFsoData.Fsos, fso_count) = FsoArray(fso_loop, 1)
            
            'Grab progress data
            If ProgressArray(fso_loop, 1) <> "" Then
                If ProgressArray(fso_loop, 1) > 0 Then
                    CollectedData(prismFsoData.Progress, fso_count) = ProgressArray(fso_loop, 1) / 100
                Else
                    CollectedData(prismFsoData.Progress, fso_count) = 0
                End If
            Else
                CollectedData(prismFsoData.Progress, fso_count) = Empty
            End If
            
            'grab and convert studio data
            If StudioArray(fso_loop, 1) = "MTL" Then
                CollectedData(prismFsoData.Studio, fso_count) = "Montréal"
            ElseIf StudioArray(fso_loop, 1) = "BUC" Then
                CollectedData(prismFsoData.Studio, fso_count) = "Bucarest"
            ElseIf StudioArray(fso_loop, 1) = "PAR" Then
                CollectedData(prismFsoData.Studio, fso_count) = "Paris"
            ElseIf StudioArray(fso_loop, 1) = "TOR" Then
                CollectedData(prismFsoData.Studio, fso_count) = "Toronto"
            ElseIf StudioArray(fso_loop, 1) = "NCT" Then
                CollectedData(prismFsoData.Studio, fso_count) = "Newcastle"
            End If

            'Grab status data
            If CollectedData(prismFsoData.Progress, fso_count) = 100 Then
                 CollectedData(prismFsoData.Status, fso_count) = "Complete"
            Else
                 CollectedData(prismFsoData.Status, fso_count) = "In Progress"
            End If
                         
            CollectedData(prismFsoData.Estimate, fso_count) = EstimateTypeArray(fso_loop, 1)
            CollectedData(prismFsoData.JIRA, fso_count) = JIRAArray(fso_loop, 1)
            CollectedData(prismFsoData.Replication, fso_count) = ReplicationArray(fso_loop, 1)
            
            'CollectedData(prismFsoData.Lines, fso_count) = TotalSPLinesArray(fso_loop, 1) + TotalMPLinesArray(fso_loop, 1)
            If TotalSPLinesArray(fso_loop, 1) = "" And TotalMPLinesArray(fso_loop, 1) = "" Then
                CollectedData(prismFsoData.Lines, fso_count) = 100
            Else
                CollectedData(prismFsoData.Lines, fso_count) = TotalSPLinesArray(fso_loop, 1) + TotalMPLinesArray(fso_loop, 1)
            End If
            
            fso_count = fso_count + 1
        End If
    Next
    
    ReDim Preserve CollectedData(1 To 8, 1 To fso_count - 1) 'TODO autocount enum total
    CollectFsoData = CollectedData
End Function


Function CollectFeatureData(targetSheet As Worksheet) As Variant

    Dim FeatureArray() As Variant
    Dim CatagoryArray() As Variant
    Dim ProgressArray() As Variant
    Dim LevelArray() As Variant
    Dim StatusArray() As Variant
    Dim StudioMTLarray() As Variant
    Dim StudioMRCarray() As Variant
    Dim StudioBUCarray() As Variant
    Dim StudioTRTarray() As Variant
    Dim StudioNCTarray() As Variant
    
    Dim HeaderMarkersArray() As Variant
    Dim feature_loop As Integer: feature_loop = 1
    Dim feature_count As Integer: feature_count = 1
    Dim CollectedData() As Variant

    'grab required date to arrays
    FeatureArray = targetSheet.Range("Table_GameFeatures[Features]").Value
    CatagoryArray = targetSheet.Range("Table_GameFeatures[Category]").Value
    ProgressArray = targetSheet.Range("Table_GameFeatures[overall_progress]").Value
    LevelArray = targetSheet.Range("Table_GameFeatures[Level]").Value
    StatusArray = targetSheet.Range("Table_GameFeatures[Feature status]").Value
    HeaderMarkersArray = targetSheet.Range("Table_GameFeatures[Status]").Value
    StudioMTLarray = targetSheet.Range("Table_GameFeatures[MTL]").Value
    StudioMRCarray = targetSheet.Range("Table_GameFeatures[MRC]").Value
    StudioBUCarray = targetSheet.Range("Table_GameFeatures[BUC]").Value
    StudioTRTarray = targetSheet.Range("Table_GameFeatures[TRT]").Value
    StudioNCTarray = targetSheet.Range("Table_GameFeatures[NCT]").Value
    
    ReDim CollectedData(1 To 7, LBound(FeatureArray) To UBound(FeatureArray))
    
    'cleans data for prism entry by:
    ' - Removing Header lines
    ' - Removing blank feature lines
    ' - Remiving Cut features
    ' - Removing dataless features (empty progression)
    
    For feature_loop = LBound(FeatureArray) To UBound(FeatureArray)
    
        If HeaderMarkersArray(feature_loop, 1) <> 1 And FeatureArray(feature_loop, 1) <> "" And StatusArray(feature_loop, 1) <> "CUT" And ProgressArray(feature_loop, 1) <> "" Then
            ' Check for missing progression data to stop results skewing
            If ProgressArray(feature_loop, 1) <> "" Then
                If ProgressArray(feature_loop, 1) > 0 Then
                    CollectedData(prismFeatureData.Complete, feature_count) = ProgressArray(feature_loop, 1) / 100
                Else
                    CollectedData(prismFeatureData.Complete, feature_count) = 0
                End If
            Else
                CollectedData(prismFeatureData.Complete, feature_count) = Empty
            End If
            
            If CollectedData(prismFeatureData.Complete, feature_count) = 1 Then
                 CollectedData(prismFeatureData.Status, feature_count) = "Complete"
            Else
                'add code to replace status with prism status
                Select Case StatusArray(feature_loop, 1)
                    Case Is = "APPROVED"
                        CollectedData(prismFeatureData.Status, feature_count) = "In Progress"
                    Case Is = "REVIEW"
                        CollectedData(prismFeatureData.Status, feature_count) = "In Progress"
                    Case Is = "SPLIT"
                        CollectedData(prismFeatureData.Status, feature_count) = "In Progress"
                    Case Is = "DETAILS NEEDED"
                        CollectedData(prismFeatureData.Status, feature_count) = "In Progress"
                End Select
            End If
                        
            CollectedData(prismFeatureData.Features, feature_count) = FeatureArray(feature_loop, 1)
            CollectedData(prismFeatureData.Catagory, feature_count) = CatagoryArray(feature_loop, 1)
            CollectedData(prismFeatureData.Level, feature_count) = "L0" 'LevelArray(feature_loop,1)
            CollectedData(prismFeatureData.Days, feature_count) = 1
            
            'grab studio in charge of feature from multile columns
            If StudioMTLarray(feature_loop, 1) = 2 Then
                CollectedData(prismFeatureData.Studio, feature_count) = "Montréal"
            ElseIf StudioBUCarray(feature_loop, 1) = 2 Then
                CollectedData(prismFeatureData.Studio, feature_count) = "Bucarest"
            ElseIf StudioMRCarray(feature_loop, 1) = 2 Then
                CollectedData(prismFeatureData.Studio, feature_count) = "Paris"
            ElseIf StudioTRTarray(feature_loop, 1) = 2 Then
                CollectedData(prismFeatureData.Studio, feature_count) = "Toronto"
            ElseIf StudioNCTarray(feature_loop, 1) = 2 Then
                CollectedData(prismFeatureData.Studio, feature_count) = "Newcastle"
            End If
            
            feature_count = feature_count + 1
        End If
    Next
    
    ReDim Preserve CollectedData(1 To 7, 1 To feature_count - 1)
    CollectFeatureData = CollectedData
    
End Function

Sub WritePrismFeatureData(PrismSheet As Worksheet, Data As Variant)
    Dim dataloop As Integer: dataloop = 1
    
    stop_excel_updates
    
    'PrismSheet.Cells(9, 1).EntireRow.Resize(UBound(Data, 2)).Insert
    For dataloop = LBound(Data, 2) To UBound(Data, 2)
        PrismSheet.Cells(7 + dataloop, prismFeatureColumns.Features).Value = Data(prismFeatureData.Features, dataloop)
        PrismSheet.Cells(7 + dataloop, prismFeatureColumns.Catagory).Value = Data(prismFeatureData.Catagory, dataloop)
        PrismSheet.Cells(7 + dataloop, prismFeatureColumns.Studio).Value = Data(prismFeatureData.Studio, dataloop)
        
        'autofill L0 as old data
        PrismSheet.Cells(7 + dataloop, prismFeatureColumns.L0_Status).Value = "Approved"
        PrismSheet.Cells(7 + dataloop, prismFeatureColumns.L0_Complete).Value = 1
        PrismSheet.Cells(7 + dataloop, prismFeatureColumns.L0_Sprint).Value = 16
        PrismSheet.Cells(7 + dataloop, prismFeatureColumns.L0_Days).Value = 1
        
        'Copy over data for L1
        PrismSheet.Cells(7 + dataloop, prismFeatureColumns.L1_Status).Value = Data(prismFeatureData.Status, dataloop)
        PrismSheet.Cells(7 + dataloop, prismFeatureColumns.L1_Complete).Value = Data(prismFeatureData.Complete, dataloop)
        PrismSheet.Cells(7 + dataloop, prismFeatureColumns.L1_Sprint).Value = "19 (Alpha)"
        PrismSheet.Cells(7 + dataloop, prismFeatureColumns.L1_Days).Value = Data(prismFeatureData.Days, dataloop)
    Next
    
    resume_excel_updates

End Sub

Sub WritePrismFSOData(PrismSheet As Worksheet, Data As Variant)
    Dim dataloop As Integer: dataloop = 1
    
    stop_excel_updates
    
    'PrismSheet.Cells(9, 1).EntireRow.Resize(UBound(Data, 2)).Insert
    For dataloop = LBound(Data, 2) To UBound(Data, 2)
        PrismSheet.Cells(7 + dataloop, prismFsoColumns.EstimateType).Value = Data(prismFsoData.Estimate, dataloop)
        PrismSheet.Cells(7 + dataloop, prismFsoColumns.FsoName).Value = Data(prismFsoData.JIRA, dataloop) & "-" & Data(prismFsoData.Fsos, dataloop)
        PrismSheet.Cells(7 + dataloop, prismFsoColumns.Studio).Value = Data(prismFsoData.Studio, dataloop)
        PrismSheet.Cells(7 + dataloop, prismFsoColumns.L1_Status).Value = Data(prismFsoData.Status, dataloop)
        'PrismSheet.Cells(7 + dataloop, prismFsoColumns.JIRA).Value = Data(prismFsoData.JIRA, dataloop)
        PrismSheet.Cells(7 + dataloop, prismFsoColumns.Replication).Value = Data(prismFsoData.Replication, dataloop)
        PrismSheet.Cells(7 + dataloop, prismFsoColumns.L1_Days).Value = Data(prismFsoData.Lines, dataloop)
        PrismSheet.Cells(7 + dataloop, prismFsoColumns.L1_Progress).Value = Data(prismFsoData.Progress, dataloop)
        PrismSheet.Cells(7 + dataloop, prismFsoColumns.L1_Sprint).Value = "19 (Alpha)"
        
    Next
    
    resume_excel_updates

End Sub

Sub ExportToPrism()
    
    Dim SourceBook As Workbook: Set SourceBook = ThisWorkbook
    Dim SourceSheet As Worksheet
    Dim DashboardSheet As Worksheet: Set DashboardSheet = SourceBook.Sheets("Dashboard")
    Dim Filename As String
    Dim Prism_sheet_URL As String
    Dim targetSheet As Worksheet
    Dim Prism_Workbook As Workbook
    Dim prismExport As Variant
    Dim prismDataRange As Range
    
    If ActiveSheet.Shapes("Check Box 1").ControlFormat.Value = -4146 And ActiveSheet.Shapes("Check Box 2").ControlFormat.Value = -4146 Then
        MsgBox ("PRISM Export Skipped - No export options were selected. Remember to set the checkboxes near the PRISM button.")
    Else
        'Open prism file and set worksheet
        If Weekday(Date, vbFriday) = 1 Then
            Filename = "DataSheet_" & Format(Date, "yyyy-mm-dd") & ".xlsm"
        Else
            Filename = "DataSheet_" & Format(Date + (8 - Weekday(Date, vbFriday)), "yyyy-mm-dd") & ".xlsm"
        End If
        
        Prism_sheet_URL = InputBox("Please check this is the correct PRISM update file, if not please correct it below:", "Export to PRISM", "\\ubisoft.org\projects\WatchDogs2\MTL\Planning\Rapports avancement\" & Filename)
        
        If Prism_sheet_URL <> "" Then
            On Error Resume Next
            Set Prism_Workbook = Workbooks.Open(Prism_sheet_URL)
            On Error GoTo 0
        
            If Prism_Workbook Is Nothing Then
                MsgBox "ERROR: Workbook " & Prism_sheet_URL & " does not exist."
            Else
                'if features export option is checked on dashboard page
                If DashboardSheet.Shapes("Check Box 1").ControlFormat.Value = 1 Then
                    Set targetSheet = Prism_Workbook.Sheets("Gameplay Features")
                    Set prismDataRange = targetSheet.Range("MTL_GPFeature_Range")
                    Set SourceSheet = SourceBook.Sheets("Game Features")
                    
                    'grab current data
                    prismExport = CollectFeatureData(SourceSheet)
                    prismDataRange.ClearContents
                    WritePrismFeatureData targetSheet, prismExport
                    MsgBox ("PRISM Feature Export Complete.")
                End If
                
                'if FSO export option is checked on dashboard page
                If DashboardSheet.Shapes("Check Box 2").ControlFormat.Value = 1 Then
                    Set targetSheet = Prism_Workbook.Sheets("FSO")
                    Set prismDataRange = targetSheet.Range("FSO_Range")
                    Set SourceSheet = SourceBook.Sheets("FSO list")
                    
                    prismExport = CollectFsoData(SourceSheet)
                    prismDataRange.ClearContents
                    WritePrismFSOData targetSheet, prismExport
                    MsgBox ("PRISM FSO Export Complete.")
                End If
            End If
        Else
            MsgBox ("PRISM Export Cancelled.")
        End If

    End If
    
End Sub
