Attribute VB_Name = "FSO_processor"
Option Explicit

Enum flColumns
    Features = 1
    fso_doc = 2
    fso_lines = 3
    fso_sp_lines = 4
    fso_spstatus_lines = 5
    fso_mp_lines = 6
    fso_mpstatus_lines = 7
End Enum

Enum fsowbColumns
    milestone = 1
    featureTag = 2
    singleP = 3
    multiP = 4
    Replication = 5
End Enum

Type config_data
    strCurrent_milestone        As String       'Current Milestone Setting e.g. ALPHA, Currently only works with one Milestone TODO
    strLine_status              As String       'FSO column name that stores Status of Line (Alpha, Gold or Cut, or maybe even levels)
    strStatus_skip              As String       'String that is flaged in the SP or MP status columns to mark FSO line as Not Applicable in those modes
    strTrack_row_ok             As String       'String that is flaged in feature_classification column to mark FSO line as OK when tested
    strTrack_row_exempt         As String       'String that is flaged in feature_classification column to mark FSO line as Cut
    intTitle_check_rows         As Integer      'Value to define max number of inital rows in which FSO column headers should lie in projects FSO files
    intTitle_check_columns      As Integer      'Value to define max number of inital columns in which FSO column headers should lie in projects FSO files
    strFSOS_URL                 As String       'String to define the root folder storing FSO files, it should include the final forward/backward slash. Works in all locations (local, Network, URL)
    strFSO_sheet_name           As String       'String to define the name of the sheet containing the key data in FSO files.
    avntApvFsoFileList()        As Variant      'Array containing a list of approved FSO file names (should not include extension, assumed to be .xlsx)
    avntFeature_list()          As Variant      'Array containing a list of all game features by name
    avntFSO_statCols_name()     As Variant      'Array of all status column names
    astrAllFsoFileList()        As Variant      'Array containing a list of all FSO file names (should not include extension, assumed to be .xlsx)
End Type

Type fso_column
    aintHeaderCell()            As Integer      'Stores the row,column cell reference for the column header
    strColumnLetter             As String       'Store Column reference as a Letter
    avntColumnData()            As Variant      'Stores the worksheet data for this column
    astrColumnName              As String       'Stores the Column Name String
    aintMsTotalLines            As Integer      'Stores the number of rows in this column for given milestone
    aintMsOkLines               As Integer      'Stores the number of rows in this column that are flagged as OK for given milestone
    
End Type

Type fso_data
    aufclProcessedFsoCols()     As fso_column   'Stores all the processed columns of hte fso, their details and data
    strProcessedFsoName         As String       'Stores the name of the processed FSO file with file extension
    intFsoLength                As Integer      'Stores the length of FSO sheet in rows *minus header rows*
End Type

Public g_wbkFeatureList         As Workbook     'The feature list excel workbook
Public g_ucdProject_Settings    As config_data  'Data variable to load all config data from sheet

Sub CONFIG_Load_fso_settings()
'
'Loads all of the FSO tools configuration data into public variable to allow all functions to access the data, and make sheet agnostic
'
    Set g_wbkFeatureList = ThisWorkbook
    
    'Clear filters in prep for reading and writing to sheet
    On Error Resume Next
    g_wbkFeatureList.Sheets("Game Features").ShowAllData
    On Error GoTo 0
    
    'bulk reading of all config data to keep it all in one place, easy to add new configs here
    With g_ucdProject_Settings
        .strCurrent_milestone = g_wbkFeatureList.Sheets("Ref Data").Range("current_milestone").Value
        .strLine_status = g_wbkFeatureList.Sheets("Ref Data").Range("line_status").Value
        .strStatus_skip = g_wbkFeatureList.Sheets("Ref Data").Range("status_skip").Value
        .strTrack_row_ok = g_wbkFeatureList.Sheets("Ref Data").Range("track_row_ok").Value
        .strTrack_row_exempt = g_wbkFeatureList.Sheets("Ref Data").Range("track_row_exempt").Value
        .intTitle_check_rows = g_wbkFeatureList.Sheets("Ref Data").Range("title_check_rows").Value
        .intTitle_check_columns = g_wbkFeatureList.Sheets("Ref Data").Range("title_check_columns").Value
        .strFSOS_URL = g_wbkFeatureList.Sheets("Ref Data").Range("FSOS_URL").Value
        .strFSO_sheet_name = g_wbkFeatureList.Sheets("Ref Data").Range("fso_sheet_name").Value
        .avntApvFsoFileList = g_wbkFeatureList.Sheets("Ref Data").Range("FSO_List").Value
        .avntFeature_list = g_wbkFeatureList.Sheets("Game Features").Range("Table_GameFeatures[Features]").Value
        .avntFSO_statCols_name = g_wbkFeatureList.Sheets("Ref Data").Range("fso_column_names").Value
        .astrAllFsoFileList = g_wbkFeatureList.Sheets("FSO list").Range(g_wbkFeatureList.Sheets("Ref Data").Range("fso_names_column").Value).Value 'TODO declair fso sheet
    End With
       
End Sub

Function Process_FSO(fso_name As String, featurelist_workbook As Workbook, aufclFSOcols() As fso_column) As fso_data
'
'opens FSO workbook, grabs data to 2d array and closes FSO workbook
'
'For returned array: TODO

    Dim wbkFSO_Workbook             As Workbook         'Stores the FSO workbook
    Dim wksFSO_Sheet                As Worksheet        'Stores the FSO data sheet
    Dim strTarget_Path              As String           'Stores path to FSO Sheet
    Dim blnFatalError               As Boolean          'Set to true if a column is missing to terminate FSO scan with error
    Dim aintTrackingColLoc(1 To 2)  As Integer          'Stores column & row number of found FSO headers
    Dim i                           As Integer          'Counter for number of FSO Columns
    Dim aufdtProcessedData          As fso_data
    
    'prevent book is already open errors making script less automated
    close_all_workbooks featurelist_workbook
    blnFatalError = False
    strTarget_Path = g_ucdProject_Settings.strFSOS_URL & fso_name 'TODO

    'check that the FSO workbook exists before trying to process it
    On Error Resume Next
    Set wbkFSO_Workbook = Workbooks.Open(strTarget_Path, ReadOnly:=True)
    On Error GoTo 0
    
    If Not wbkFSO_Workbook Is Nothing Then
    
        'Check worksheet names, expecting name from worksheet config
        On Error Resume Next
        Set wksFSO_Sheet = wbkFSO_Workbook.Sheets(g_ucdProject_Settings.strFSO_sheet_name) 'TODO
        On Error GoTo 0
        
        If Not wksFSO_Sheet Is Nothing Then
        
            'clear filters to prevent not reading hidden items
            On Error Resume Next
            wksFSO_Sheet.ShowAllData
            On Error GoTo 0
            
            'Find each column in FSO sheet, check and output errors if not present
            blnFatalError = findColumnHeaders(aufclFSOcols, aintTrackingColLoc, wksFSO_Sheet, fso_name)

            'If fatal errors with columns are found, we dont read the data from this FSO Workbook to not pollute sheet data until it's fixed.
            If blnFatalError = False Then
                'Otherwise collect the data into fso data type for later processing
                aufdtProcessedData.intFsoLength = wksFSO_Sheet.Cells(wksFSO_Sheet.Cells.Rows.Count, ConvertToLetter(aintTrackingColLoc(cellLocation.Column))).End(xlUp).Row - aintTrackingColLoc(cellLocation.Row)
                aufdtProcessedData.strProcessedFsoName = fso_name
                For i = 1 To UBound(aufclFSOcols)
                    aufclFSOcols(i).strColumnLetter = ConvertToLetter(aufclFSOcols(i).aintHeaderCell(cellLocation.Column))
                    aufclFSOcols(i).avntColumnData = wksFSO_Sheet.Range(aufclFSOcols(i).strColumnLetter & (aintTrackingColLoc(cellLocation.Row) + 1) & ":" & aufclFSOcols(i).strColumnLetter & (aufdtProcessedData.intFsoLength + aintTrackingColLoc(cellLocation.Row)))
                Next
                
                countFSOStatusLines aufclFSOcols, aufdtProcessedData.intFsoLength
                aufdtProcessedData.aufclProcessedFsoCols = aufclFSOcols 'TODO could be just one variable
                               
                VLS_WriteLogging LoggingSheet, wbkFSO_Workbook.name & ": Workbook was read successfully.", VLS_ErrorTypes.success
            End If
           
            'close open FSO workbook
            wbkFSO_Workbook.Close False
        Else
            'close open FSO workbook
            wbkFSO_Workbook.Close False
            VLS_WriteLogging LoggingSheet, fso_name & ": A sheet with the name " & g_ucdProject_Settings.strFSO_sheet_name & " was not found. Workbook was skipped", VLS_ErrorTypes.Error, VLS_ErrorTags.workbook_error
        End If
    Else
        VLS_WriteLogging LoggingSheet, fso_name & ": Workbook did not exist. Workbook was skipped", VLS_ErrorTypes.Error, VLS_ErrorTags.workbook_error
    End If
    Process_FSO = aufdtProcessedData
End Function

Function findColumnHeaders(aufclFsoColumns() As fso_column, aintHeaderLoc() As Integer, wksFSO As Worksheet, name As String) As Boolean
'
' Finds headers in FSO files matching passed coumn names and sets column & row number of the header, If header isnt found returns true. else false
'
    Dim i As Integer
    Dim j As Integer
    
    findColumnHeaders = False
    
    'Find each column in FSO sheet, check and output errors if not present
    For i = 1 To UBound(aufclFsoColumns)
        aufclFsoColumns(i).aintHeaderCell = find_column_number(wksFSO, g_ucdProject_Settings.intTitle_check_rows, g_ucdProject_Settings.intTitle_check_columns, aufclFsoColumns(i).astrColumnName)
        'Need to store the title row if a header is found. Allows calculation of total number of data rows in sheet
        If aufclFsoColumns(i).aintHeaderCell(cellLocation.Row) > -1 Then
            aintHeaderLoc(cellLocation.Row) = aufclFsoColumns(i).aintHeaderCell(cellLocation.Row)
        End If
        
        'If this FSO has column errors, write errors out to log
        If (aufclFsoColumns(i).aintHeaderCell(cellLocation.Column) <= 0) Then
            findColumnHeaders = True
            VLS_WriteLogging LoggingSheet, name & ": Workbook is missing key column " & aufclFsoColumns(i).astrColumnName & ". Workbook was skipped", VLS_ErrorTypes.Error, VLS_ErrorTags.column_error
        Else
            'We use the fso line status column to dictate a line we need to track in our stats, so storing this column number specfically to use as reference
            If aufclFsoColumns(i).astrColumnName = g_ucdProject_Settings.strLine_status Then
                aintHeaderLoc(cellLocation.Column) = aufclFsoColumns(i).aintHeaderCell(cellLocation.Column)
            End If
        End If
    Next
End Function

Sub countFSOStatusLines(aufclFsoColumns() As fso_column, intTotalFSOLines As Integer)
'
'Fuction takes fso column data and counts total number of lines and OK statuses
'
    Dim i As Integer
    Dim j As Integer

    'TODO magic number, 3 is the first status column to be read from config so better to start at 3. Need a neater way to do this
    For i = 3 To UBound(aufclFsoColumns)
        'initalise values to not inherit values from previous fso aufclFsoColumns is persistent through reading of all FSO files
        aufclFsoColumns(i).aintMsTotalLines = 0
        aufclFsoColumns(i).aintMsOkLines = 0
        
        For j = 1 To intTotalFSOLines
        
            'check if the milstone set in config matches the milestone flagged in fso line
            If str_SafeComp(g_ucdProject_Settings.strCurrent_milestone) = str_SafeComp(aufclFsoColumns(1).avntColumnData(j, 1)) Then
            
                'Count total lines
                aufclFsoColumns(i).aintMsTotalLines = aufclFsoColumns(i).aintMsTotalLines + 1
                'Remove any lines flagged as Cut
                If str_SafeComp(aufclFsoColumns(i).avntColumnData(j, 1)) = str_SafeComp(g_ucdProject_Settings.strTrack_row_exempt) Then
                    aufclFsoColumns(i).aintMsTotalLines = aufclFsoColumns(i).aintMsTotalLines - 1
                End If
                'Remove any lines flagged as Not Applicable
                If str_SafeComp(aufclFsoColumns(i).avntColumnData(j, 1)) = str_SafeComp(g_ucdProject_Settings.strStatus_skip) Then
                    aufclFsoColumns(i).aintMsTotalLines = aufclFsoColumns(i).aintMsTotalLines - 1
                End If
                'COunt the line if its status is Okay
                If str_SafeComp(aufclFsoColumns(i).avntColumnData(j, 1)) = str_SafeComp(g_ucdProject_Settings.strTrack_row_ok) Then
                    aufclFsoColumns(i).aintMsOkLines = aufclFsoColumns(i).aintMsOkLines + 1
                End If
            
            End If
            
        Next
        
    Next
                
End Sub
Sub calculateFeatureStats(fl_workbook As Workbook, fso_name As String, processed_FSO() As Variant, flwb_data() As Variant)
'
'calculates all of the FSO stats and puts them into array
'
    'loop counters
    Dim feature_loop        As Integer:     feature_loop = 1
    Dim fso_loop            As Integer:     fso_loop = 1
    'Line matching counters
    Dim valid_lines         As Integer:     valid_lines = 0
    Dim matched_lines       As Integer:     matched_lines = 0
    Dim fso_first_pass      As Boolean:     fso_first_pass = False
    Dim fso_missing_tags    As Integer:     fso_missing_tags = 0
    
    Dim milestone           As String:      milestone = g_ucdProject_Settings.strCurrent_milestone
    
    'check for stats logic between feature list and FSO
    For feature_loop = LBound(flwb_data(flColumns.Features), 1) To UBound(flwb_data(flColumns.Features), 1)
        For fso_loop = LBound(processed_FSO, 1) To UBound(processed_FSO, 1)
            
            'System for scanning for missed feature labels in FSO files. Check for any Non-Cut lines missing feature tags and counts them for later warnings
            If str_SafeComp(processed_FSO(fso_loop, fsowbColumns.milestone)) <> g_ucdProject_Settings.strTrack_row_exempt And processed_FSO(fso_loop, fsowbColumns.milestone) <> Empty And _
                fso_first_pass = False And processed_FSO(fso_loop, fsowbColumns.featureTag) = Empty Then
                
                fso_missing_tags = fso_missing_tags + 1
            End If
            
            
            'if feature names matches tag in FSO and featurelist entry isnt not blank and line is not marked as cut
            If str_SafeComp(flwb_data(flColumns.Features)(feature_loop, 1)) = str_SafeComp(processed_FSO(fso_loop, 2)) _
                And str_SafeComp(processed_FSO(fso_loop, 1)) <> g_ucdProject_Settings.strTrack_row_exempt And flwb_data(flColumns.Features)(feature_loop, 1) <> Empty Then
                    
                'Total FSO Lines
                If processed_FSO(fso_loop, 1) <> "" Then
                    flwb_data(flColumns.fso_lines)(feature_loop, 1) = flwb_data(flColumns.fso_lines)(feature_loop, 1) + 1
                End If
                    
                'Total SP lines
                If str_SafeComp(processed_FSO(fso_loop, 1)) = milestone And str_SafeComp(processed_FSO(fso_loop, 3)) <> g_ucdProject_Settings.strStatus_skip Then
                    flwb_data(flColumns.fso_sp_lines)(feature_loop, 1) = flwb_data(flColumns.fso_sp_lines)(feature_loop, 1) + 1
                End If
                    
                'Total MP lines
                If str_SafeComp(processed_FSO(fso_loop, 1)) = milestone And str_SafeComp(processed_FSO(fso_loop, 4)) <> g_ucdProject_Settings.strStatus_skip Then
                    flwb_data(flColumns.fso_mp_lines)(feature_loop, 1) = flwb_data(flColumns.fso_mp_lines)(feature_loop, 1) + 1
                End If
                
'                'Total Replication lines TODO
'                If str_SafeComp(processed_FSO(FSO_loop, 1)) = milestone And str_SafeComp(processed_FSO(FSO_loop, 4)) <> g_ucdProject_Settings.strStatus_skip Then
'                    flwb_data(flColumns.fso_mp_lines)(feature_loop, 1) = flwb_data(flColumns.fso_mp_lines)(feature_loop, 1) + 1
'                End If
'
                'Total OKAY SP status lines
                If str_SafeComp(processed_FSO(fso_loop, 3)) = g_ucdProject_Settings.strTrack_row_ok And str_SafeComp(processed_FSO(fso_loop, 1)) = milestone Then 'TODO
                    flwb_data(flColumns.fso_spstatus_lines)(feature_loop, 1) = flwb_data(flColumns.fso_spstatus_lines)(feature_loop, 1) + 1
                End If
                   
                'Total OKAY MP status lines
                If str_SafeComp(processed_FSO(fso_loop, 4)) = g_ucdProject_Settings.strTrack_row_ok And str_SafeComp(processed_FSO(fso_loop, 1)) = milestone Then 'TODO
                    flwb_data(flColumns.fso_mpstatus_lines)(feature_loop, 1) = flwb_data(flColumns.fso_mpstatus_lines)(feature_loop, 1) + 1
                End If
                
                'store FSO names separated by comma
                If flwb_data(flColumns.fso_doc)(feature_loop, 1) = Empty Then
                    flwb_data(flColumns.fso_doc)(feature_loop, 1) = Replace(fso_name, ".xlsx", "")
                ElseIf InStr(flwb_data(flColumns.fso_doc)(feature_loop, 1), Replace(fso_name, ".xlsx", "")) = 0 Then
                    flwb_data(flColumns.fso_doc)(feature_loop, 1) = flwb_data(flColumns.fso_doc)(feature_loop, 1) & Chr(10) & Replace(fso_name, ".xlsx", "")
                End If
                
                'Set empty values that should have data to Zero for better presentation in sheet.
                If flwb_data(flColumns.fso_lines)(feature_loop, 1) = Empty Then flwb_data(flColumns.fso_lines)(feature_loop, 1) = 0
                If flwb_data(flColumns.fso_sp_lines)(feature_loop, 1) = Empty Then flwb_data(flColumns.fso_sp_lines)(feature_loop, 1) = 0
                If flwb_data(flColumns.fso_mp_lines)(feature_loop, 1) = Empty Then flwb_data(flColumns.fso_mp_lines)(feature_loop, 1) = 0
                If flwb_data(flColumns.fso_spstatus_lines)(feature_loop, 1) = Empty Then flwb_data(flColumns.fso_spstatus_lines)(feature_loop, 1) = 0
                If flwb_data(flColumns.fso_mpstatus_lines)(feature_loop, 1) = Empty Then flwb_data(flColumns.fso_mpstatus_lines)(feature_loop, 1) = 0

            End If
        Next
        
        'mark first pass off FSO complete to avoid double counting line errors
        fso_first_pass = True
        
    Next
    
    'write to log if any lines are missing feature tags
    If fso_missing_tags > 0 Then
        VLS_WriteLogging LoggingSheet, fso_name & ": FSO has " & fso_missing_tags & " line(s) with missing feature tags.", VLS_ErrorTypes.warning, VLS_ErrorTags.line_error
    End If
    
End Sub

Sub createJaggedFeaturelistArray(avntJaggedArray As Variant)
'
'Create a set of jagged arrays to store the featurelist stats the processing will calculate so they can be easily written back to a table range
'
    'create arrays for Arraythe key feature list worksheet columns
    Dim avntJaggedDimensions()      As Variant
    Dim intFlLowerBound             As Integer: intFlLowerBound = LBound(g_ucdProject_Settings.avntFeature_list, 1)
    Dim intFlUpperBound             As Integer: intFlUpperBound = UBound(g_ucdProject_Settings.avntFeature_list, 1)
    
    'set up empty arrays to match feature list dimensions
    ReDim avntJaggedArray(1 To 7) 'TODO this should NOT be a magic number 7
    ReDim avntJaggedDimensions(intFlLowerBound To intFlUpperBound, 1 To 1)

    'create jagged array with all feature list workbook data
    avntJaggedArray(flColumns.Features) = g_ucdProject_Settings.avntFeature_list
    avntJaggedArray(flColumns.fso_doc) = avntJaggedDimensions
    avntJaggedArray(flColumns.fso_lines) = avntJaggedDimensions
    avntJaggedArray(flColumns.fso_sp_lines) = avntJaggedDimensions
    avntJaggedArray(flColumns.fso_spstatus_lines) = avntJaggedDimensions
    avntJaggedArray(flColumns.fso_mp_lines) = avntJaggedDimensions
    avntJaggedArray(flColumns.fso_mpstatus_lines) = avntJaggedDimensions

End Sub


Sub writeFeatureData(avntFlData As Variant)
'
'Write calculated feature stats data back to featurelist. Source is x by 1 array hence direct writes to table column ranges
'
    With g_wbkFeatureList.Sheets("Game Features")
        .Range("Table_GameFeatures[fso_doc]") = avntFlData(flColumns.fso_doc)
        .Range("Table_GameFeatures[fso_lines]") = avntFlData(flColumns.fso_lines)
        .Range("Table_GameFeatures[fso_sp_lines]") = avntFlData(flColumns.fso_sp_lines)
        .Range("Table_GameFeatures[fso_spstatus_lines]") = avntFlData(flColumns.fso_spstatus_lines)
        .Range("Table_GameFeatures[fso_mp_lines]") = avntFlData(flColumns.fso_mp_lines)
        .Range("Table_GameFeatures[fso_mpstatus_lines]") = avntFlData(flColumns.fso_mpstatus_lines)
    End With
End Sub

Sub initialise_columns(avntColumns() As fso_column)
'
' Reads array of status column names from config variable into a column type
'
    
    Dim cintNames       As Integer      'Counter for the number of status columns
    Dim intColumnNum    As Integer      'Number of satus columns names read from config
    
    intColumnNum = UBound(g_ucdProject_Settings.avntFSO_statCols_name)
    ReDim avntColumns(1 To intColumnNum)
    
    For cintNames = 1 To intColumnNum
        avntColumns(cintNames).astrColumnName = g_ucdProject_Settings.avntFSO_statCols_name(cintNames, 1)
    Next
    
End Sub

Sub writeFsoData(aufdtFsoDatatoWrite() As fso_data)
    
    Dim intNumFsos           As Integer: intNumFsos = UBound(g_ucdProject_Settings.astrAllFsoFileList)
    'Dim aintFsoSPData()      As Integer
    'Dim aintFsoMPData()      As Integer
    'Dim aintFsoRepData()     As Integer
    Dim i                    As Integer

    ReDim aintFsoSPData(1 To intNumFsos, 1 To 1)
    ReDim aintFsoSPTotalData(1 To intNumFsos, 1 To 1)
    ReDim aintFsoMPData(1 To intNumFsos, 1 To 1)
    ReDim aintFsoMPTotalData(1 To intNumFsos, 1 To 1)
    ReDim aintFsoRepData(1 To intNumFsos, 1 To 1)
    ReDim aintFsoRepTotalData(1 To intNumFsos, 1 To 1)

    '1 To intNumFsos, 1 To 1
    For i = 1 To intNumFsos
        If str_SafeComp(aufdtFsoDatatoWrite(i).strProcessedFsoName) = str_SafeComp(g_ucdProject_Settings.astrAllFsoFileList(i, 1) & ".xlsx") And aufdtFsoDatatoWrite(i).strProcessedFsoName <> "" Then
            
            If aufdtFsoDatatoWrite(i).aufclProcessedFsoCols(3).aintMsTotalLines > 0 Then
                aintFsoSPData(i, 1) = aufdtFsoDatatoWrite(i).aufclProcessedFsoCols(3).aintMsOkLines
                aintFsoSPTotalData(i, 1) = aufdtFsoDatatoWrite(i).aufclProcessedFsoCols(3).aintMsTotalLines
            Else
                aintFsoSPData(i, 1) = 0
                aintFsoSPTotalData(i, 1) = 0
            End If
            
            If aufdtFsoDatatoWrite(i).aufclProcessedFsoCols(4).aintMsTotalLines > 0 Then
                aintFsoMPData(i, 1) = aufdtFsoDatatoWrite(i).aufclProcessedFsoCols(4).aintMsOkLines
                aintFsoMPTotalData(i, 1) = aufdtFsoDatatoWrite(i).aufclProcessedFsoCols(4).aintMsTotalLines
            Else
                aintFsoMPData(i, 1) = 0
                aintFsoMPTotalData(i, 1) = 0
            End If
            
            If aufdtFsoDatatoWrite(i).aufclProcessedFsoCols(5).aintMsTotalLines > 0 Then
                aintFsoRepData(i, 1) = aufdtFsoDatatoWrite(i).aufclProcessedFsoCols(5).aintMsOkLines
                aintFsoRepTotalData(i, 1) = aufdtFsoDatatoWrite(i).aufclProcessedFsoCols(5).aintMsTotalLines
            Else
                aintFsoRepData(i, 1) = 0
                aintFsoRepTotalData(i, 1) = 0
            End If
    
        End If
    Next
    
    'Clear filters in prep for reading and writing to sheet
    On Error Resume Next
    g_wbkFeatureList.Sheets("FSO list").ShowAllData
    On Error GoTo 0
    
    g_wbkFeatureList.Sheets("FSO list").Range("Table_FSOList[SP FSO Status]").Value = aintFsoSPData
    g_wbkFeatureList.Sheets("FSO list").Range("Table_FSOList[SP FSO Total]").Value = aintFsoSPTotalData
    g_wbkFeatureList.Sheets("FSO list").Range("Table_FSOList[MP FSO Status]").Value = aintFsoMPData
    g_wbkFeatureList.Sheets("FSO list").Range("Table_FSOList[MP FSO Total]").Value = aintFsoMPTotalData
    g_wbkFeatureList.Sheets("FSO list").Range("Table_FSOList[Replication FSO Status]").Value = aintFsoRepData
    g_wbkFeatureList.Sheets("FSO list").Range("Table_FSOList[Replication FSO Total]").Value = aintFsoRepTotalData
     
End Sub

Sub refresh_fso_stats()
'
'MAIN FUCTION FOR FSO REFRESH
'
    Dim fso_data()                  As Variant      'Store required data from FSO files
    Dim aufdtProcessedFsoData()     As fso_data
    Dim avntFsoFileList()           As Variant      'List of all FSO files to be processed
    Dim data_check                  As Integer      'Stores size of FSO_Data after process to be sure its not empty
    Dim fso_loop                    As Integer      'Loop for FSO files
    Dim strTargetFSOname            As String       'Current FSO file to be processed
    Dim avntFlWorkbookData()        As Variant      'Feature Stats eventually written back to feature list sheet
    Dim aufclFsoColumns()           As fso_column   'Column data
    Dim i, j                        As Integer
    
    CONFIG_Load_fso_settings 'load global config from excel sheet cells for project fsos
    initialise_columns aufclFsoColumns() 'reads config data into column data type
    ReDim aufdtProcessedFsoData(1 To UBound(g_ucdProject_Settings.astrAllFsoFileList)) 'create fso type per fso
    
    
    avntFsoFileList = g_ucdProject_Settings.avntApvFsoFileList 'Grab list of FSO files to process
    createJaggedFeaturelistArray avntFlWorkbookData 'Create jagged array to store all data. 2d arrays with x by 1 write direct to table ranges without loops.
    VLS_initialise_logging 1 'LOGGING SYSTEM: setup and clear the logging

    For fso_loop = LBound(avntFsoFileList, 1) To UBound(avntFsoFileList, 1)
    
        'check if the cell the value was read from didnt contain an error (was happening alot in the spreadsheet cell, and checks its not empty TODO
        If VarType(avntFsoFileList(fso_loop, 1)) <> vbError Then
            If avntFsoFileList(fso_loop, 1) <> Empty Then
        
                'Load and process all FSOs
                strTargetFSOname = avntFsoFileList(fso_loop, 1) & ".xlsx"
                aufdtProcessedFsoData(fso_loop) = Process_FSO(strTargetFSOname, g_wbkFeatureList, aufclFsoColumns()) 'fso_data = Process_FSO(strTargetFSOname, g_wbkFeatureList, aufclFsoColumns())
                
                'Check for the return of an empty FSO array (i.e a read/processing error happened)
                If aufdtProcessedFsoData(fso_loop).strProcessedFsoName = "" Then 'FSO was read incorrectly
                    VLS_WriteLogging LoggingSheet, strTargetFSOname & ": Workbook stats writing failed.", VLS_ErrorTypes.Error
                Else 'FSO was read correctly, crunch the data to add to the featurelist
                    
                    '**********TEMP ADAPTOR TODO **********
                    ReDim fso_data(1 To aufdtProcessedFsoData(fso_loop).intFsoLength, 1 To UBound(aufdtProcessedFsoData(fso_loop).aufclProcessedFsoCols))
                    For i = 1 To UBound(aufdtProcessedFsoData(fso_loop).aufclProcessedFsoCols)
                        For j = 1 To (aufdtProcessedFsoData(fso_loop).intFsoLength)
                            'dump column data into array
                            fso_data(j, i) = aufdtProcessedFsoData(fso_loop).aufclProcessedFsoCols(i).avntColumnData(j, 1) 'TODO write whole column data here instead?
                        Next
                    Next
                    '**********TEMP ADAPTOR TODO **********
                    
                    calculateFeatureStats g_wbkFeatureList, strTargetFSOname, fso_data, avntFlWorkbookData
                    VLS_WriteLogging LoggingSheet, strTargetFSOname & ": Workbook stats writing successful.", VLS_ErrorTypes.success
                End If
            End If
    
        Else
            VLS_WriteLogging LoggingSheet, "FSO List Item " & fso_loop & ": Cell contains a reference error, please check it.", VLS_ErrorTypes.Error, VLS_ErrorTags.system_error
        End If
    
    Next

    'write arrays back to worksheet & write log entry
    writeFsoData aufdtProcessedFsoData()
    writeFeatureData (avntFlWorkbookData)
    VLS_WriteLogging LoggingSheet, "FSO Data Refresh Complete", VLS_ErrorTypes.success
    
    'log date & time of last refresh
    If check_range_exists("last_refresh_date", g_wbkFeatureList.Sheets("Ref Data")) = True Then g_wbkFeatureList.Sheets("Ref Data").Range("last_refresh_date").Value = Format(Now(), "DD-MM-YY")
    If check_range_exists("last_refresh_time", g_wbkFeatureList.Sheets("Ref Data")) = True Then g_wbkFeatureList.Sheets("Ref Data").Range("last_refresh_time").Value = Format(Now(), "HH:MM:SS")
    
    MsgBox "FSO Stats Refresh Complete!"

End Sub




