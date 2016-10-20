Attribute VB_Name = "Basics_Library"
'ENUMS for check_file_exists() Function
Enum filePathType
    LocalFile = 1
    NetworkFile = 2
    WebFile = 3
End Enum

Enum cellLocation
    Row = 1
    Column = 2
End Enum



Sub close_all_workbooks(exception As Workbook)

    Dim wkb As Workbook
    
    For Each wkb In Workbooks
        If exception.name <> wkb.name Then
            wkb.Close False
        End If
    Next

End Sub

Function ConvertToLetter(iCol As Integer) As String
'
'COnverts a Column Number to a Column Letter. Returns String Value, Accepts Column Number
'
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      ConvertToLetter = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
   End If
End Function

Function str_SafeComp(dangerString As Variant)
'
'fixes common typing errors the cause strings to fail comparison checks
'
    str_SafeComp = Replace(UCase(dangerString), " ", "")
    str_SafeComp = Replace(str_SafeComp, "-", "")

End Function

Sub stop_excel_updates()
'
'Stops key excel updating functions to improve sheet writing speeds
'
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

End Sub

Sub resume_excel_updates()
'
'Resumes key excel updating functions to improve sheet writing speeds
'
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

End Sub

Function check_file_exists(FilePath As String, Optional pathType As Integer) As Integer
'
'Function to check if a file exists, returns, True, False or Error if string is blank
'Auto-detects if file is local, on network or sharepoint based on passsed URL OR
'you can force it with the optional path type (see ENUMS at top of file)
'
    Dim URLTypeOK   As Boolean: URLTypeOK = False 'Checks that the url type has been detected, to provide error if it passes through fuction and is not.
    
    If Len(FilePath) = 0 Then
        check_file_exists = False
        Exit Function
    End If
    
    'check file is on local file path
    If Mid(FilePath, 2, 2) = ":\" Or pathType = filePathType.LocalFile Then
        URLTypeOK = True
        If Len(Dir(FilePath)) > 0 Then
            check_file_exists = True
        Else
            check_file_exists = False
        End If
    End If
    
    'check if file is on network area
    If Mid(FilePath, 1, 2) = "\\" Or pathType = filePathType.NetworkFile Then
        URLTypeOK = True
        If Len(Dir(FilePath)) > 0 Then
            check_file_exists = True
        Else
            check_file_exists = False
        End If
    End If
    
    'check if file is on sharepoint
    'formula =IF([@Filename]<>"",IF(check_file_exists(FSOS_URL & D2 & ".xlsx"),HYPERLINK(FSOS_URL & D2 & ".xlsx","[EXCEL]"),"MISSING"),"ADD FILE NAME")
    
    If Mid(FilePath, 1, 7) = "http://" Or pathType = filePathType.WebFile Or Mid(FilePath, 1, 8) = "https://" Then
        URLTypeOK = True
        Dim oHttpRequest As Object
        Set oHttpRequest = New MSXML2.XMLHTTP60
        
        On Error Resume Next
            With oHttpRequest
                .Open "GET", FilePath, False
                .setRequestHeader "Cache-Control", "no-cache"
                .setRequestHeader "Pragma", "no-cache"
                .send
            End With
        On Error GoTo 0

        If oHttpRequest.Status = 200 Then
            check_file_exists = True
        Else
            check_file_exists = False
        End If
    End If
    
    If URLTypeOK = False Then
        MsgBox ("URL Validation Error: Url type is not recognised. Check the Url prefix and that this fuction supports it.")
    End If
    
End Function

Function find_column_number(targetSheet As Worksheet, maxRow As Integer, maxCol As Integer, header As String) As Integer()
'
'Function that retunrs row and column number of found headers otherwise returns both values as -1
'Retuns 2 value array
'

            Dim rowCol(1 To 2) As Integer
            Dim i As Integer
            Dim j As Integer
            rowCol(1) = -1
            rowCol(2) = -1
            
            'loop for the first few rows and columns to track down the  eaders
            For i = 1 To maxRow
                For j = 1 To maxCol
                    If str_SafeComp(targetSheet.Cells(i, j).Value) = str_SafeComp(header) Then
                        rowCol(cellLocation.Column) = j
                        rowCol(cellLocation.Row) = i
                        Exit For
                    End If
                Next
                If rowCol(cellLocation.Column) > -1 Then
                    Exit For
                End If
            Next
            
            find_column_number = rowCol
End Function

Function check_range_exists(strRange As String, wksRangeSheet As Worksheet) As Boolean
'
' Function to check if named range exists
'
    Dim rngRangeCheck As Range
    
    On Error Resume Next
    Set rngRangeCheck = wksRangeSheet.Range(strRange)
    On Error GoTo 0
    
    If rngRangeCheck Is Nothing Then
       check_range_exists = False
    Else
       check_range_exists = True
    End If

End Function
