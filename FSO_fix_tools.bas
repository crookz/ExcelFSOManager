Attribute VB_Name = "FSO_fix_tools"
Option Explicit

Sub ApplyFSOconditionalFormatting()
'
' Resets conditional formatting inside of FSO
'
'
    Dim spColumn        As String
    Dim mpColumn        As String
    Dim repColumn       As String
    Dim msColumn        As String
    
    Dim applyRange      As Range
    Dim milestoneRange  As Range
    Dim referenceRange  As String
    
    'not all fso formats are the same!! Get user to double check the column letters and correct them if wrong
    spColumn = InputBox(Prompt:="Enter the column letter of SP Status", Title:="Status Colmun Info", Default:="G")
    mpColumn = InputBox(Prompt:="Enter the column letter of MP Status", Title:="Status Colmun Info", Default:="J")
    repColumn = InputBox(Prompt:="Enter the column letter of Replication Status", Title:="Status Colmun Info", Default:="M")
    msColumn = InputBox(Prompt:="Enter the column letter of Dev Priority ", Title:="Status Colmun Info", Default:="B")
    
          
    Set applyRange = ActiveSheet.Range(spColumn & "5:" & spColumn & "500," & mpColumn & "5:" & mpColumn & "500," & repColumn & "5:" & repColumn & "500")
    
    ActiveSheet.Cells.FormatConditions.Delete
    
    'APPLY STATUS CELLS--------------------------------------------------------------------------
        'OKAY TEST
        applyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="Okay"
        With applyRange.FormatConditions(applyRange.FormatConditions.Count).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 5287936 'piercing green
            .TintAndShade = 0
        End With
        applyRange.FormatConditions(applyRange.FormatConditions.Count).StopIfTrue = False
        
        'NOT OKAY TEST
        applyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="Not Okay"
        With applyRange.FormatConditions(applyRange.FormatConditions.Count).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
        End With
        applyRange.FormatConditions(applyRange.FormatConditions.Count).StopIfTrue = False
        
        'NOT IMPLEMENTED TEST
        applyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="Not Implemented"
        With applyRange.FormatConditions(applyRange.FormatConditions.Count).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 10498160
            .TintAndShade = 0
        End With
        applyRange.FormatConditions(applyRange.FormatConditions.Count).StopIfTrue = False
         
        'NA TEST
        applyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="NA"
        With applyRange.FormatConditions(applyRange.FormatConditions.Count).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.249977111117893
        End With
        applyRange.FormatConditions(applyRange.FormatConditions.Count).StopIfTrue = False
                 
        'PENDING GD TEST
        applyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="Pending GD"
        With applyRange.FormatConditions(applyRange.FormatConditions.Count).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0
        End With
        applyRange.FormatConditions(applyRange.FormatConditions.Count).StopIfTrue = False
        
        'In Progress Test
        applyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="In Progress"
        With applyRange.FormatConditions(applyRange.FormatConditions.Count).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0
        End With
        applyRange.FormatConditions(applyRange.FormatConditions.Count).StopIfTrue = False
        
        'TBT TEST
        applyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="TBT"
        With applyRange.FormatConditions(applyRange.FormatConditions.Count).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 5296274
            .TintAndShade = 0
        End With
        applyRange.FormatConditions(applyRange.FormatConditions.Count).StopIfTrue = False
        
        'Flag missing status on gold lines
        Set applyRange = ActiveSheet.Range(spColumn & "5:" & spColumn & "500")
        applyRange.FormatConditions.Add Type:=xlExpression, Formula1:="=AND($B5=""Gold"",ISBLANK(" & spColumn & "5))"
        With applyRange.FormatConditions(applyRange.FormatConditions.Count)
        
            .Borders(xlLeft).LineStyle = xlDot
            .Borders(xlLeft).TintAndShade = 0
            .Borders(xlLeft).Weight = xlThin
            
            .Borders(xlRight).LineStyle = xlDot
            .Borders(xlRight).TintAndShade = 0
            .Borders(xlRight).Weight = xlThin
            
            .Borders(xlTop).LineStyle = xlDot
            .Borders(xlTop).TintAndShade = 0
            .Borders(xlTop).Weight = xlThin
            
            .Borders(xlBottom).LineStyle = xlDot
            .Borders(xlBottom).TintAndShade = 0
            .Borders(xlBottom).Weight = xlThin

        End With
        applyRange.FormatConditions(applyRange.FormatConditions.Count).StopIfTrue = False
        
        Set applyRange = ActiveSheet.Range(mpColumn & "5:" & mpColumn & "500")
        applyRange.FormatConditions.Add Type:=xlExpression, Formula1:="=AND($B5=""Gold"",ISBLANK(" & mpColumn & "5))"
        With applyRange.FormatConditions(applyRange.FormatConditions.Count)
        
            .Borders(xlLeft).LineStyle = xlDot
            .Borders(xlLeft).TintAndShade = 0
            .Borders(xlLeft).Weight = xlThin
            
            .Borders(xlRight).LineStyle = xlDot
            .Borders(xlRight).TintAndShade = 0
            .Borders(xlRight).Weight = xlThin
            
            .Borders(xlTop).LineStyle = xlDot
            .Borders(xlTop).TintAndShade = 0
            .Borders(xlTop).Weight = xlThin
            
            .Borders(xlBottom).LineStyle = xlDot
            .Borders(xlBottom).TintAndShade = 0
            .Borders(xlBottom).Weight = xlThin

        End With
        applyRange.FormatConditions(applyRange.FormatConditions.Count).StopIfTrue = False
        
        Set applyRange = ActiveSheet.Range(repColumn & "5:" & repColumn & "500")
        applyRange.FormatConditions.Add Type:=xlExpression, Formula1:="=AND($B5=""Gold"",ISBLANK(" & repColumn & "5))"
        With applyRange.FormatConditions(applyRange.FormatConditions.Count)
        
            .Borders(xlLeft).LineStyle = xlDot
            .Borders(xlLeft).TintAndShade = 0
            .Borders(xlLeft).Weight = xlThin
            
            .Borders(xlRight).LineStyle = xlDot
            .Borders(xlRight).TintAndShade = 0
            .Borders(xlRight).Weight = xlThin
            
            .Borders(xlTop).LineStyle = xlDot
            .Borders(xlTop).TintAndShade = 0
            .Borders(xlTop).Weight = xlThin
            
            .Borders(xlBottom).LineStyle = xlDot
            .Borders(xlBottom).TintAndShade = 0
            .Borders(xlBottom).Weight = xlThin

        End With
        applyRange.FormatConditions(applyRange.FormatConditions.Count).StopIfTrue = False
        
        
    'APPLY EMPTY STATUS CELLS--------------------------------------------------------------------------
        Set applyRange = ActiveSheet.Range(spColumn & "5:" & spColumn & "500," & mpColumn & "5:" & mpColumn & "500," & repColumn & "5:" & repColumn & "500")
        
        'Flag empty milestone/team cells test
        Set milestoneRange = Range("B5:B500")
        milestoneRange.FormatConditions.Add Type:=xlExpression, Formula1:="=AND(COUNTA($C5)=1,COUNTA($B5)=0)"
        With milestoneRange.FormatConditions(milestoneRange.FormatConditions.Count).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.399975585192419
        End With
        milestoneRange.FormatConditions(milestoneRange.FormatConditions.Count).StopIfTrue = False
        
        'Flag empty milestone/team cells test
        Set milestoneRange = Range(msColumn & "5:" & msColumn & "500")
        milestoneRange.FormatConditions.Add Type:=xlExpression, Formula1:="=AND(COUNTA($B5)=1,COUNTA($C5)=0)"
        With milestoneRange.FormatConditions(milestoneRange.FormatConditions.Count).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.399975585192419
        End With
        milestoneRange.FormatConditions(milestoneRange.FormatConditions.Count).StopIfTrue = False

    'APPLY ALTERNATE BACKDROP STATUS CELLS--------------------------------------------------------------------------
        'Flag empty status cells (alternating colours) test
        referenceRange = "$B$5:$B$500"
        applyRange.FormatConditions.Add Type:=xlExpression, Formula1:="=AND(ROW()>2,COUNTA($B5)>0,MOD(ROW(),2)=0)"
        With applyRange.FormatConditions(applyRange.FormatConditions.Count).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 12443629
            .TintAndShade = 0
        End With
        applyRange.FormatConditions(applyRange.FormatConditions.Count).StopIfTrue = False
        
        applyRange.FormatConditions.Add Type:=xlExpression, Formula1:="=AND(ROW()>2,COUNTA($B5)>0,MOD(ROW(),2)=1)"
        With applyRange.FormatConditions(applyRange.FormatConditions.Count).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 8438492
            .TintAndShade = 0
        End With
        applyRange.FormatConditions(applyRange.FormatConditions.Count).StopIfTrue = False

    
End Sub

Sub validate_FSO_links()
'
'Tool to check all FSO document links and update their status in the FSO list tab.
'Will using MISSING if file is not found, or add hyperlink if file is present
'
    Dim FSOfiles        As Range: Set FSOfiles = ActiveSheet.Range("Table_FSOList[Summary]")
    Dim FSOLink         As Range: Set FSOLink = ActiveSheet.Range("Table_FSOList[SharePoint]")
    Dim FSOjiraURLs     As Range: Set FSOjiraURLs = ActiveSheet.Range("Table_FSOList[FSO Link]")

    Dim CellLoop        As Integer
    Dim FSOurl          As String
    
    CONFIG_Load_fso_settings
    stop_excel_updates
    
    For CellLoop = 1 To FSOfiles.Rows.Count
        'clear existing hyperlinks
        FSOLink.Cells(CellLoop, 1).Hyperlinks.Delete
        
        FSOurl = FSOjiraURLs.Cells(CellLoop, 1).Value
        If check_file_exists(FSOurl) Then
            ActiveSheet.Hyperlinks.Add _
            Anchor:=FSOLink.Cells(CellLoop, 1), Address:=FSOurl, ScreenTip:="Excel File: " & FSOurl, TextToDisplay:="[OPEN]"
        Else
            FSOLink.Cells(CellLoop, 1).Value = "MISSING"
        End If
        
    Next CellLoop
    resume_excel_updates
    MsgBox ("FSO File Check Complete")
End Sub




