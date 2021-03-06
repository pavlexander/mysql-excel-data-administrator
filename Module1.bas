Attribute VB_Name = "Module1"
Sub generateFile()

Dim ws As Worksheet
Dim ws_main As Worksheet: Set ws_main = ActiveWorkbook.Worksheets("main")
Dim WS_Count As Integer
Dim I As Integer
Dim Column_count As String
Dim Row_Count As String
Dim insertValues As String
Dim separator As String: separator = ","
Dim cellValue As String
Dim tableName As String
Dim insertCommand As String

Dim OutputFileNum As Integer
Dim PathName As String

Dim FileName As String
Dim FileExtension As String
Dim useStatement As String
Dim TablesTotal As Integer
Dim InsertsTotal As Integer


ws_main.Range("TBL_TOT").Value = ""
ws_main.Range("INS_TOT").Value = ""

FileName = ws_main.Range("FILE_NAME") 'Cells(2, 5)
FileExtension = ws_main.Range("FILE_EXT") 'Cells(3, 5)
useStatement = ws_main.Range("USE_SQL") 'Cells(4, 5)

WS_Count = ActiveWorkbook.Worksheets.Count

TablesTotal = WS_Count - 1

If WS_Count > 1 Then

    PathName = Application.ActiveWorkbook.Path
    OutputFileNum = FreeFile
    
    Open PathName & "\" & FileName & "." & FileExtension For Output Lock Write As #OutputFileNum

    For I = 1 To WS_Count
       ' MsgBox ActiveWorkbook.Worksheets(I).Name
        Set ws = ActiveWorkbook.Worksheets(I)
       
        If ws.Name <> "main" Then
            tableName = ws.Name
            'MsgBox tableName
            
            Column_count = ws.UsedRange.Columns.Count
            Row_Count = ws.UsedRange.Rows.Count
            
            If Row_Count > 3 And Column_count > 2 Then
                
                'MsgBox Column_count
                'MsgBox Row_Count
                
                For row = 4 To Row_Count
                    insertValues = ""
                    
                    For col = 1 To Column_count
                        If col = 1 Then
                            If ws.Cells(row, col) = "Ignore Row" Then
                                Exit For 'don't itterate columns of this row
                            Else
                                GoTo ContinueLoop 'skip first column
                            End If
                        End If
                    
                        'MsgBox ws.Cells(row, col).Value
                    
                        If ws.Cells(row, col) = "" Then
                            If ws.Cells(2, col) = "" Then
                                Exit For
                            ElseIf ws.Cells(2, col) = "DEFAULT" Then
                                cellValue = "DEFAULT"
                                
                                insertValues = insertValues & (separator & cellValue)
                            ElseIf ws.Cells(2, col) = "NULL" Then
                                cellValue = ""
                                
                                insertValues = insertValues & (separator & cellValue)
                            Else
                                cellValue = ws.Cells(2, col).Value
                                
                                insertValues = insertValues & (separator & cellValue)
                            End If
                        Else
                            If ws.Cells(1, col) <> "" Then
                                If ws.Cells(1, col) = "NUMBER" Then
                                    cellValue = ws.Cells(row, col).Value
                                End If
                            Else
                                cellValue = ws.Cells(row, col).Value
                                
                                ' escape single quote
                                cellValue = Replace(cellValue, "'", "\'")
                                
                                ' escape double quote
                                cellValue = Replace(cellValue, """", "\""")
                                
                                ' enclose result value into single quotes
                                cellValue = "'" & cellValue & "'"
                                
                            End If
                            
                            insertValues = insertValues & (separator & cellValue)
                        End If
                        
ContinueLoop:
                    Next col
                    
                    If Len(insertValues) <> 0 Then
                        insertValues = Right$(insertValues, (Len(insertValues) - Len(separator)))
                    End If
                    
                    If insertValues <> "" Then
                        InsertsTotal = InsertsTotal + 1
                    
                        'MsgBox insertValues
                        If useStatement = "Yes" Then
                            insertCommand = "INSERT INTO {tableName} VALUES ({insertValues});"
                            insertCommand = Replace(insertCommand, "{tableName}", tableName)
                            insertCommand = Replace(insertCommand, "{insertValues}", insertValues)
                            Print #OutputFileNum, insertCommand
                        Else
                            Print #OutputFileNum, insertValues
                        End If
                    End If

                Next row
            End If
        End If
    Next I
    
    Close OutputFileNum

End If

ws_main.Range("TBL_TOT").Value = TablesTotal
ws_main.Range("INS_TOT").Value = InsertsTotal
      
End Sub

Sub AddWSTable()

Dim ws As Worksheet
Dim ws_main As Worksheet: Set ws_main = ActiveWorkbook.Worksheets("main")
Dim insertLine As String
Dim openPos As Integer
Dim closePos As Integer
Dim midBit As String
Dim WrdArray() As String
Dim headerCellValue As String
Dim lastColumn As Integer

Dim matchesFound As Collection
Dim tableName As String

insertLine = ws_main.Range("INS_STMT").Value '14, 4

If insertLine <> "" Then
    Set matchesFound = getSeparatedValues(insertLine, "`")
    'MsgBox matchesFound.Count
    'MsgBox matchesFound(1)
    tableName = matchesFound(2)
    
    If SheetExists(tableName) = True Then
        MsgBox "Error. Worksheet (table) with name '" & tableName & "' already exists."
    Else
        openPos = InStr(insertLine, "(")
        closePos = InStr(insertLine, ")")
        midBit = Mid(insertLine, openPos + 1, closePos - openPos - 1)
        
        'MsgBox midBit
        
        WrdArray() = Split(midBit, ",")
        
        If UBound(WrdArray) > 0 Then
            Set ws = ThisWorkbook.Sheets.Add(After:= _
                     ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            ws.Name = tableName
            
            For I = LBound(WrdArray) To UBound(WrdArray)
                headerCellValue = WrdArray(I)
                headerCellValue = Trim(headerCellValue)
                headerCellValue = Replace(headerCellValue, "`", "")
                ws.Cells(3, I + 2).Value = headerCellValue
                
                If headerCellValue = "id" Then
                    ws.Cells(1, I + 2).Value = "NUMBER"
                ElseIf EndsWith(headerCellValue, "_by") Then
                    ws.Cells(1, I + 2).Value = "NUMBER"
                ElseIf EndsWith(headerCellValue, "_id") Then
                    ws.Cells(1, I + 2).Value = "NUMBER"
                End If
                
                ' add validations
                ws.Cells(1, I + 2).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="NUMBER"
                ws.Cells(1, I + 2).Validation.ErrorMessage = "Please select a valid value from the list"
                ws.Cells(2, I + 2).Validation.Add Type:=xlValidateList, Formula1:="DEFAULT,NULL,1,0"
                ws.Cells(2, I + 2).Validation.ShowError = False
                
                lastColumn = I + 2
                
                ws.Cells(1, I + 2).EntireColumn.AutoFit
                ws.Cells(1, I + 2).EntireColumn.HorizontalAlignment = xlCenter
            Next I
            
            ' set first column values
            ws.Cells(1, 1).Value = ws_main.Range("ROW_TYPE").Value
            ws.Cells(2, 1).Value = ws_main.Range("DEFAULT_VALUE").Value
            ws.Cells(3, 1).Value = ws_main.Range("COLUMN_NAME").Value
            'ws.Cells(4, 1).Value = ws_main.Range("DATA_ROWS").Value
            
            ' set colors
            ws.Cells(1, 1).EntireRow.Interior.Color = ws_main.Range("COLOR1").Interior.Color '16, 4
            ws.Cells(2, 1).EntireRow.Interior.Color = ws_main.Range("COLOR2").Interior.Color '17, 4
            ws.Cells(3, 1).EntireRow.Interior.Color = ws_main.Range("COLOR3").Interior.Color '18, 4
            
            ' set column after which safe to enter any text
            ws.Cells(1, lastColumn + 1).EntireColumn.Interior.Color = ws_main.Range("AFTER_LAST_COL").Interior.Color
            
            ' and additional parameters
            ws.Cells(1, 1).EntireColumn.AutoFit
            ws.Cells(1, 1).EntireColumn.HorizontalAlignment = xlRight
            ws.Cells(1, 1).EntireColumn.Interior.Color = ws_main.Range("ROW_TYPE").Interior.Color
            
            ' set borders
            ws.Cells(3, 1).EntireRow.Borders(xlEdgeBottom).LineStyle = xlContinuous
            ws.Cells(1, 1).EntireColumn.Borders(xlEdgeRight).LineStyle = xlContinuous
            
            'set validation for first column
            ws.Cells(1, 1).EntireColumn.Validation.Add Type:=xlValidateList, Formula1:="Ignore Row"
            ws.Cells(1, 1).Validation.Delete
            ws.Cells(2, 1).Validation.Delete
            ws.Cells(3, 1).Validation.Delete
            
            'set conditional formatting rule for rows which will be ignored
            With ws.Range("=$A$1:$Z$1500")
                .FormatConditions.Add Type:=xlExpression, Formula1:="=INDIRECT(""a""&ROW())=""Ignore Row"""
                .FormatConditions(.FormatConditions.Count).Interior.Color = ws_main.Range("IGNORE_ROW").Interior.Color
            End With
        End If
    End If
End If

End Sub

Private Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

     If wb Is Nothing Then Set wb = ThisWorkbook
     On Error Resume Next
     Set sht = wb.Sheets(shtName)
     On Error GoTo 0
     SheetExists = Not sht Is Nothing
 End Function
 
Private Function getSeparatedValues(sText As String, char As String) As Collection
    Dim getSeparatedValues_ As New Collection
    
    Dim bIsBetween As Boolean
    Dim skipNext As Boolean
    
    Dim iLength As Integer
    
    Dim sToken As String
    
    bIsBetween = False
    skipNext = False
    
    sToken = ""
    
    iLength = Len(sText) - 1
    
    For I = 1 To iLength
        If (skipNext = True) Then
            skipNext = False
        Else
            Dim chr As String
            Dim nextChr As String
        
            chr = Mid(sText, I, 1)
            nextChr = Mid(sText, I + 1, 1)
        
            If (chr = char) Then
                bIsBetween = True
            End If
        
            If (nextChr = char) Then
                bIsBetween = False
            End If
        
            If (bIsBetween = True) Then
                sToken = sToken & nextChr
            Else
                If (Len(sToken) > 0) Then
                    skipNext = True
                    getSeparatedValues_.Add (sToken)
                    sToken = ""
                End If
            End If
        End If
    Next I

    Set getSeparatedValues = getSeparatedValues_
    Set getSeparatedValues_ = Nothing
End Function

Private Function EndsWith(str As String, ending As String) As Boolean
     Dim endingLen As Integer
     endingLen = Len(ending)
     EndsWith = (Right(Trim(UCase(str)), endingLen) = UCase(ending))
End Function

Private Function StartsWith(str As String, start As String) As Boolean
     Dim startLen As Integer
     startLen = Len(start)
     StartsWith = (Left(Trim(UCase(str)), startLen) = UCase(start))
End Function

