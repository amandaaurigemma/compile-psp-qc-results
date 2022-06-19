'VB macro for mac users to compile PSP QC metric data
'Written by Amanda Aurigemma
'Version 1.0.0 || Last updated 29DEC2021

Sub requestFileAccess()
    
    Dim myFilePath As String, filesInPath As String, myFiles() As String
    Dim i As Integer

'Establish path and request folder access
    myFilePath = Worksheets("Macro Settings").Range("B1").Text

    filesInPath = Dir(myFilePath)
    If filesInPath = "" Then
        MsgBox "No files found in specified folder"
        Exit Sub
    End If
    
'Create file directory
    i = 0
    Do While filesInPath <> ""
        i = i + 1
        ReDim Preserve myFiles(1 To i)
        myFiles(i) = filesInPath
        filesInPath = Dir()
    Loop
    
'Print list of files to new sheet
    ActiveWorkbook.Worksheets.Add
    If i > 0 Then
        For i = LBound(myFiles) To UBound(myFiles)
            On Error Resume Next
            Cells(i, 2).Value = myFiles(i)
        Next i
    End If
    
'Add path and concat
    ActiveCell.FormulaR1C1 = "='Macro Settings'!R1C2"
    Range("A1").Select
    Selection.AutoFill Destination:=Range("A1:A" & i - 1)
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "=CONCAT(RC[-2]:RC[-1])"
    Range("C1").Select
    Selection.AutoFill Destination:=Range("C1:C" & i - 1)
    Range("C1:C" & i - 1).Select
    Selection.Copy
    Range("D1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'Create new array
    Dim filePermissionCandidates As Variant
    
    filePermissionCandidates = Application.Transpose(Range("D1:D" & i - 1))
    
'Delete sheet
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True
    
'Grant access to files
    Dim fileAccessGranted As Boolean
    fileAccessGranted = GrantAccessToMultipleFiles(filePermissionCandidates)

End Sub

Sub getAllData()
    
    requestFileAccess
    
    Dim myFilePath As String, filesInPath As String, myFiles() As String
    Dim i As Integer
    Dim outputSheet As Worksheet
    
    myFilePath = Worksheets("Macro Settings").Range("B1").Text
    Set outputSheet = Worksheets(Worksheets("Macro Settings").Range("B2").Value)
    
    filesInPath = Dir(myFilePath)
    If filesInPath = "" Then
        MsgBox "No files found in specified folder"
        Exit Sub
    End If
    
    i = 0
    Do While filesInPath <> ""
        i = i + 1
        ReDim Preserve myFiles(1 To i)
        myFiles(i) = filesInPath
        filesInPath = Dir()
    Loop
    
    'Pull sample names from file names
    If i > 0 Then
        For i = LBound(myFiles) To UBound(myFiles)
        On Error Resume Next
            With outputSheet
                 .Cells(i + 1, 1).Value = Left(myFiles(i), Worksheets("Macro Settings").Range("B3").Value)
                 Call getFileData(myFiles(i), i + 1, Worksheets("Macro Settings").Range("B2").Value)
            End With
        Next i
    End If
    
    End Sub
   
Sub getFileData(fileName As String, rowNum As Integer, wrksht As String)

'Establish file path

Dim filePath As String
filePath = Worksheets("Macro Settings").Range("B1").Text & fileName

'Open new sheet
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets.Add

'Import metric Data from specified JSON file (configured using record macro)
    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=Range("$A$1"))
        .Name = fileName
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .RefreshPeriod = False
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 10000
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
'Move metrics to output sheet
    For i = 30 To 100
        If Range("C" + CStr(i)).Value = "Per-Sample Read Count" Then
            Worksheets(wrksht).Range("C" + CStr(rowNum)).Value = Range("C" + CStr(i + 1)).Value
            
        ElseIf Range("C" + CStr(i)).Value = "Total Primer Count" Then
            Worksheets(wrksht).Range("D" + CStr(rowNum)).Value = Range("C" + CStr(i + 1)).Value
            
        ElseIf Range("C" + CStr(i)).Value = "Percent of Primers with a Read Count > 10" Then
            Worksheets(wrksht).Range("E" + CStr(rowNum)).Value = Range("C" + CStr(i + 1)).Value
            
        ElseIf Range("C" + CStr(i)).Value = "Percent of Raw Reads Containing Expected Primer Sequence" Then
            Worksheets(wrksht).Range("F" + CStr(rowNum)).Value = Range("C" + CStr(i + 1)).Value
            
        ElseIf Range("C" + CStr(i)).Value = "Median PSP Read Count Divided by Median SNP ID Read Count" Then
            Worksheets(wrksht).Range("G" + CStr(rowNum)).Value = Range("C" + CStr(i + 1)).Value
        
        ElseIf Range("C" + CStr(i)).Value = "Mean PSP Read Count Divided by Mean SNP ID Read Count" Then
            Worksheets(wrksht).Range("H" + CStr(rowNum)).Value = Range("C" + CStr(i + 1)).Value
        
        ElseIf Range("C" + CStr(i)).Value = "Synthesis Control Uniformity" Then
            Worksheets(wrksht).Range("I" + CStr(rowNum)).Value = Range("C" + CStr(i + 1)).Value
        
        ElseIf Range("C" + CStr(i)).Value = "PSP Uniformity" Then
            Worksheets(wrksht).Range("J" + CStr(rowNum)).Value = Range("C" + CStr(i + 1)).Value
        
        ElseIf Range("C" + CStr(i)).Value = "Number of Primers With Less Than 10 Raw Reads" Then
            Worksheets(wrksht).Range("K" + CStr(rowNum)).Value = Range("C" + CStr(i + 1)).Value
        
        ElseIf Range("C" + CStr(i)).Value = "Total Number of Reads Assigned to an Unexpected Primer" Then
            Worksheets(wrksht).Range("L" + CStr(rowNum)).Value = Range("C" + CStr(i + 1)).Value
        
        ElseIf Range("C" + CStr(i)).Value = "Number of Unknown Sequences With Greater Than 10 Raw Reads" Then
            Worksheets(wrksht).Range("M" + CStr(rowNum)).Value = Range("C" + CStr(i + 1)).Value
        
        ElseIf Range("C" + CStr(i)).Value = "Total Number of Unknown Sequences" Then
            Worksheets(wrksht).Range("N" + CStr(rowNum)).Value = Range("C" + CStr(i + 1)).Value
        
        ElseIf Range("C" + CStr(i)).Value = "Adapter Dimer Count" Then
            Worksheets(wrksht).Range("O" + CStr(rowNum)).Value = Range("C" + CStr(i + 1)).Value
        
        ElseIf Range("C" + CStr(i)).Value = "Percent of Reads Assigned to Expected Target Primers or Known Adapter Dimer Sequences" Then
            Worksheets(wrksht).Range("P" + CStr(rowNum)).Value = Range("C" + CStr(i + 1)).Value
        
        ElseIf Range("C" + CStr(i)).Value = "Number of Duplicate Targeted Primer Sequences" Then
            Worksheets(wrksht).Range("Q" + CStr(rowNum)).Value = Range("C" + CStr(i + 1)).Value
            
        ElseIf Range("C" + CStr(i)).Value = "Number of Total Raw Reads Pre-Subsample" Then
            Worksheets(wrksht).Range("R" + CStr(rowNum)).Value = Range("C" + CStr(i + 1)).Value
            
        End If
    Next
    
 'Move date to output sheet
    For i = 10 To 20
        If Range("C" + CStr(i)).Value = "De-multiplexing Date" Then
            Worksheets(wrksht).Range("T" + CStr(rowNum)).Value = Left(Range("C" + CStr(i + 1)).Value, 4) + "-" + Mid(Range("C" + CStr(i + 1)).Value, 5, 2) + "-" + Mid(Range("C" + CStr(i + 1)).Value, 7, 2)
        End If
    Next
    
 'Move config to output sheet
    Dim endRow As Integer
    endRow = Range("B2").End(xlDown).Row
    For i = endRow - 10 To endRow
        If Range("B" + CStr(i)).Value = "configuration:" Then
            Worksheets(wrksht).Range("S" + CStr(rowNum)).Value = Range("C" + CStr(i)).Value
    End If
    Next
    
 'Close the data import worksheet
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True
    
 'Close the data Import Query
    ActiveWorkbook.Queries(fileName).Delete
    
    RemoveConnections
    
End Sub
Sub RemoveConnections()
    Dim conn As Long
    With ActiveWorkbook
        For conn = .Connections.Count To 1 Step -1
            .Connections(conn).Delete
        Next conn
    End With
End Sub
