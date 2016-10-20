Attribute VB_Name = "Logging_System"
'**********************VBA LOGGING SYSTEM*******************************************
'Author: Sean Crooks
'Uses a sheet called Log to write out logging lines.
'Auto create log sheet for you if missing
'
'USAGE:
'   - Declare public variabe at top of main module: Public LoggingSheet As Worksheet
'   - Call initialisation module: initialise_logging
'   - custoimise ENUM tag names below & in write function
'Setup is complete, Call functions when needed to write to or clear log
'
'***********************************************************************************

Option Explicit

Enum VLS_ErrorTags
    system_error = 1 '**RESERVED DO NOT CHANGE**
    line_error = 2
    column_error = 3
    workbook_error = 4
    custom5_error = 5
End Enum

Enum VLS_ErrorTypes
    success = 0
    Error = 1
    warning = 2
End Enum

Public LoggingSheet             As Worksheet    'Defined for the logging System

Sub VLS_initialise_logging(newLog As Integer)
'
'Sets up intital logging file, checking for existence and recreation if nessesary
'if newLog = 1 the log is cleared
'
    Set LoggingSheet = Nothing
    
    On Error Resume Next
    Set LoggingSheet = ThisWorkbook.Sheets("Log")
    On Error GoTo 0
    
    If LoggingSheet Is Nothing Then
        'create a new log sheet and add warning that it was missing
        ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).name = "Log"
        Set LoggingSheet = ThisWorkbook.Sheets("Log")
        LoggingSheet.Range("A1").FormulaR1C1 = "Date"
        LoggingSheet.Range("B1").FormulaR1C1 = "Time"
        LoggingSheet.Range("C1").FormulaR1C1 = "Description"
        LoggingSheet.Range("D1").FormulaR1C1 = "Type"
        LoggingSheet.Range("E1").FormulaR1C1 = "Tag"
        LoggingSheet.Columns("A:A").ColumnWidth = 10
        LoggingSheet.Columns("B:B").ColumnWidth = 10
        LoggingSheet.Columns("C:C").ColumnWidth = 100
        LoggingSheet.Columns("D:D").ColumnWidth = 15
        LoggingSheet.Columns("E:E").ColumnWidth = 15
        LoggingSheet.Range("A1:E1").AutoFilter
        'LoggingSheet.Visible = xlSheetHidden
        
        VLS_WriteLogging LoggingSheet, "LOGGING SYSTEM: Log sheet did not exist, a new one was created.", VLS_ErrorTypes.warning, 1
    Else
        If newLog = 1 Then
            VLS_clear_log LoggingSheet
        End If
    End If
End Sub

Sub VLS_clear_log(Log As Worksheet)
'
'Clears all logging
'
    Dim last_entry As Long
    
    'clear filters to prevent read last line code not reading hidden items
    On Error Resume Next
    LoggingSheet.ShowAllData
    On Error GoTo 0

    last_entry = Log.Cells(Log.Rows.Count, "A").End(xlUp).Row + 1
    Log.Rows("2:" & last_entry).EntireRow.Delete
End Sub

Sub VLS_WriteLogging(Log As Worksheet, Entry As String, ErrorType As Integer, Optional ErrorTag As Integer)
'
'Writes Logging to log sheet. ErrorType 0=Success, 1= Error, 2= Warning
'
    Dim last_entry As Long
    
    'CUSTOMISE YOUR TAGS HERE
    Dim ErrorTag1 As String: ErrorTag1 = "SYSTEM" '**RESERVED DO NOT CHANGE**
    Dim ErrorTag2 As String: ErrorTag2 = "LINE"
    Dim ErrorTag3 As String: ErrorTag3 = "COLUMN"
    Dim ErrorTag4 As String: ErrorTag4 = "WORKBOOK"
    Dim ErrorTag5 As String: ErrorTag5 = "CUSTOM TAG 5"
    
    If ErrorType > 2 Or ErrorType < 0 Then
        ErrorType = 1
        'get last entry ro insert new entry after
        last_entry = Log.Cells(Log.Rows.Count, "A").End(xlUp).Row + 1
       
        'Write Error Log
        Log.Cells(last_entry, "A").Value = Format(Now(), "DD-MM-YY")
        Log.Cells(last_entry, "B").Value = Format(Now(), "HH:MM:SS")
        Log.Cells(last_entry, "C").Value = "LOGGING SYSTEM: Incorrect error type passed to logging fuction for" & ": " & Entry
       
        Select Case ErrorType
            Case 0
                Log.Cells(last_entry, "D").Value = "SUCCESS"
            Case 1
                Log.Cells(last_entry, "D").Value = "ERROR"
            Case 2
                Log.Cells(last_entry, "D").Value = "WARNING"
        End Select
        
         Log.Cells(last_entry, "E").Value = "SYSTEM"
    Else
       'get last entry row insert new entry after
       last_entry = Log.Cells(Log.Rows.Count, "A").End(xlUp).Row + 1
       
       'Write Error Log
       Log.Cells(last_entry, "A").Value = Format(Now(), "DD-MM-YY")
       Log.Cells(last_entry, "B").Value = Format(Now(), "HH:MM:SS")
       Log.Cells(last_entry, "C").Value = Entry
       
       Select Case ErrorType
        Case 0
            Log.Cells(last_entry, "D").Value = "SUCCESS"
        Case 1
            Log.Cells(last_entry, "D").Value = "ERROR"
        Case 2
            Log.Cells(last_entry, "D").Value = "WARNING"
       End Select
       
       If ErrorTag > 0 Then
        Select Case ErrorTag
            Case 1
                Log.Cells(last_entry, "E").Value = ErrorTag1
            Case 2
                Log.Cells(last_entry, "E").Value = ErrorTag2
            Case 3
                Log.Cells(last_entry, "E").Value = ErrorTag3
            Case 4
                Log.Cells(last_entry, "E").Value = ErrorTag4
            Case 5
                Log.Cells(last_entry, "E").Value = ErrorTag5
        End Select
       End If
    End If

End Sub

