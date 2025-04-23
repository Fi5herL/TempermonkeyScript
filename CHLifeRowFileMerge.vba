Option Explicit ' 強制宣告所有變數

Sub Button1_Click()
    Call MergeSpecificExcelFiles_And_Process_V3
End Sub

Sub MergeSpecificExcelFiles_And_Process_V3() ' 再改個版本號

    Dim fso As Object ' FileSystemObject
    Dim sourceFolder As Object
    Dim sourceFile As Object
    Dim folderPath As String
    Dim destWb As Workbook
    Dim sourceWb As Workbook
    Dim sourceSheet As Worksheet
    Dim newSheetName As String
    Dim fileCount As Long
    Dim initialSheetCount As Integer ' 記錄新活頁簿初始的工作表數量

    ' --- 1. 讓使用者選擇來源資料夾 ---
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇包含來源檔案的資料夾"
        .AllowMultiSelect = False
        If .Show <> -1 Then ' 如果使用者取消選擇
            MsgBox "操作已取消。", vbInformation
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With

    ' --- 2. 初始化 ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        MsgBox "選擇的資料夾不存在：" & folderPath, vbCritical
        GoTo CleanUp
    End If
    Set sourceFolder = fso.GetFolder(folderPath)

    ' --- 3. 建立新的目標活頁簿 ---
    Set destWb = Workbooks.Add
    initialSheetCount = destWb.Sheets.Count
    fileCount = 0

    ' --- 4. 遍歷資料夾中的檔案 (合併階段) ---
    On Error Resume Next ' 暫時忽略錯誤
    For Each sourceFile In sourceFolder.Files
        Err.Clear
        Dim fileName As String
        Dim fileExt As String
        fileName = LCase(sourceFile.Name)
        fileExt = LCase(fso.GetExtensionName(fileName))

        ' --- 5. 檢查檔案是否符合條件 ---
        If (Left(fileName, 6) = "export" Or Left(fileName, 5) = "聚會資料_") And _
           (fileExt = "xlsx" Or fileExt = "xls" Or fileExt = "xlsm" Or fileExt = "csv") Then

            ' --- 6. 開啟來源檔案並複製工作表 ---
            Set sourceWb = Workbooks.Open(sourceFile.Path, ReadOnly:=False)

            If Err.Number = 0 Then
                If sourceWb.Sheets.Count > 0 Then
                    Set sourceSheet = sourceWb.Sheets(1)
                    newSheetName = fso.GetBaseName(sourceFile.Name)
                    newSheetName = CleanSheetName(newSheetName)
                    sourceSheet.Copy After:=destWb.Sheets(destWb.Sheets.Count)
                    destWb.Sheets(destWb.Sheets.Count).Name = newSheetName
                    fileCount = fileCount + 1
                End If
                sourceWb.Close SaveChanges:=False
            Else
                Debug.Print "無法開啟檔案: " & sourceFile.Path & " - 錯誤: " & Err.Description
                Err.Clear
            End If
            Set sourceSheet = Nothing
            Set sourceWb = Nothing
        End If
    Next sourceFile
    On Error GoTo 0 ' 恢復正常錯誤處理

    ' --- 7. 清理目標活頁簿的預設工作表 ---
    If destWb.Sheets.Count > initialSheetCount Then
        Dim i As Integer
        For i = initialSheetCount To 1 Step -1
             destWb.Sheets(i).Delete
        Next i
    ElseIf fileCount = 0 Then
        destWb.Close SaveChanges:=False
        MsgBox "在指定的資料夾中沒有找到符合條件的檔案。", vbInformation
        GoTo CleanUp
    End If

    ' --- 8a. 後處理：為 Export 表格添加 "會所名稱" ---
    Dim ws As Worksheet
    Dim lookupSheet As Worksheet
    Dim exportSheet As Worksheet
    Dim origFileNameCol As Variant
    Dim huosuoNameCol As Variant
    Dim foundRow As Variant
    Dim huosuoName As String
    Dim lastRow As Long ' Reused later
    Dim foundLookupSheet As Boolean

    foundLookupSheet = False
    Set lookupSheet = Nothing

    ' 先找到有效的 "聚會資料_" 工作表
    For Each ws In destWb.Worksheets
        If LCase(ws.Name) Like "聚會資料_*" Then
            On Error Resume Next ' Prevent error if Match fails
            origFileNameCol = Application.Match("原始檔案名稱", ws.Rows(1), 0)
            huosuoNameCol = Application.Match("會所名稱", ws.Rows(1), 0)
            On Error GoTo 0 ' Restore error handling

            If Not IsError(origFileNameCol) And Not IsError(huosuoNameCol) Then
                Set lookupSheet = ws
                foundLookupSheet = True
                Debug.Print "找到用於查找的 '聚會資料_' 工作表: " & lookupSheet.Name
                Exit For
            Else
                 Debug.Print "工作表 '" & ws.Name & "' 開頭為 '聚會資料_' 但缺少必要標頭。"
                 origFileNameCol = Empty ' Reset if only one was found
                 huosuoNameCol = Empty
            End If
        End If
    Next ws

    ' 填充 "會所名稱"
    If foundLookupSheet Then
        For Each exportSheet In destWb.Worksheets
            If LCase(exportSheet.Name) Like "export*" Then
                Debug.Print "處理 Export 工作表 (填充會所): " & exportSheet.Name
                On Error Resume Next ' Handle potential error during Match
                foundRow = Application.Match(exportSheet.Name, lookupSheet.Columns(origFileNameCol), 0)
                On Error GoTo 0

                If Not IsError(foundRow) Then
                    huosuoName = CStr(lookupSheet.Cells(foundRow, huosuoNameCol).Value)
                    Debug.Print "  找到匹配。會所名稱: " & huosuoName

                    exportSheet.Columns(1).Insert Shift:=xlToRight
                    exportSheet.Cells(1, 1).Value = "會所名稱"

                    ' Find last row based on original Column A (now Column B)
                    If exportSheet.Cells(exportSheet.Rows.Count, "B").End(xlUp).Row = 1 And IsEmpty(exportSheet.Cells(1, "B").Value) Then
                        lastRow = 1
                    Else
                        lastRow = exportSheet.Cells(exportSheet.Rows.Count, "B").End(xlUp).Row
                    End If

                    ' Fill down starting from A2
                    If lastRow >= 2 Then
                        exportSheet.Range("A2:A" & lastRow).Value = ConvertToHallNumber(huosuoName)
                        Debug.Print "    已填入 A2:A" & lastRow
                    ElseIf lastRow = 1 Then
                         exportSheet.Range("A2").Value = ConvertToHallNumber(huosuoName) ' Fill A2 even if only header exists in B
                         Debug.Print "    僅填入 A2。"
                    ' Else: lastRow = 0 (empty sheet), do nothing
                    End If
                Else
                    Debug.Print "  未找到匹配的原始檔案名稱: '" & exportSheet.Name & "'"
                    exportSheet.Columns(1).Insert Shift:=xlToRight
                    exportSheet.Cells(1, 1).Value = "會所名稱(未找到)"
                End If
            End If
        Next exportSheet
    Else
        MsgBox "警告: 未找到有效的 '聚會資料_' 查找表，無法填充 '會所名稱'。", vbExclamation
    End If


' --- 8b. 修改後邏輯：重新命名第一個，合併其餘相同 B1 的 ---
    Dim targetSheet As Worksheet      ' 作為目標的工作表 (被改名的那個)
    Dim secondarySheet As Worksheet ' 作為數據來源的工作表
    Dim b1Value As String           ' 當前處理的 B1 分組值
    Dim targetSheetName As String   ' 目標工作表的名稱 (可能改變)
    Dim processedSheetNames As Object ' 用於記錄已處理的工作表名，避免重複作為 target
    Dim sourceLastRow As Long
    Dim sourceLastCol As Long
    Dim destNextRow As Long
    Dim wsIndex As Integer          ' 用於遍歷工作表的索引
    Dim secondaryWsIndex As Integer ' 用於內部遍歷的索引

    Set processedSheetNames = CreateObject("Scripting.Dictionary") ' 用字典記錄 B1 值是否已被處理過

    Debug.Print "----- 開始 8b：重新命名與合併 Export 工作表 -----"

    ' 外層循環：找到下一個未處理的 Export 表作為 Target
    For wsIndex = 1 To destWb.Worksheets.Count
        On Error Resume Next ' 以防工作表在處理過程中被意外刪除
        Set targetSheet = Nothing
        Set targetSheet = destWb.Worksheets(wsIndex)
        If Err.Number <> 0 Then
            Err.Clear
            Debug.Print "  跳過索引 " & wsIndex & " (可能已刪除)"
            GoTo NextOuterSheet ' 跳到下一個外層循環
        End If
        On Error GoTo 0

        ' 檢查是否是 Export 表 且 B1 尚未處理 (B1 有內容)
        If LCase(targetSheet.Name) Like "export*" And Len(Trim(CStr(targetSheet.Range("B1").Value))) > 0 Then

            b1Value = Trim(CStr(targetSheet.Range("B1").Value))
            targetSheetName = targetSheet.Name ' 記錄當前名稱

            ' 檢查這個 B1 值是否已經被處理過了（即是否有其他表因為這個 B1 值被改名了）
            If Not processedSheetNames.Exists(b1Value) Then
                ' --- 這是此 B1 值的第一個目標 ---
                Debug.Print "  找到新的目標分組 B1: '" & b1Value & "', 來源工作表: '" & targetSheetName & "'"

                ' 1. 重新命名目標工作表
                Dim newNameCandidate As String
                newNameCandidate = CleanSheetName(b1Value)
                If Len(newNameCandidate) > 0 And LCase(newNameCandidate) <> LCase(targetSheetName) Then
                    On Error Resume Next ' 捕獲命名衝突
                    targetSheet.Name = newNameCandidate
                    If Err.Number <> 0 Then
                        Debug.Print "    **警告**: 無法將工作表 '" & targetSheetName & "' 重新命名為 '" & newNameCandidate & "'. Error: " & Err.Description & ". 將維持原名處理。"
                        Err.Clear
                        ' 如果重命名失敗，仍然使用原名 targetSheetName 作為目標
                    Else
                        targetSheetName = newNameCandidate ' 更新目標名稱
                        Debug.Print "    工作表已重新命名為: '" & targetSheetName & "'"
                    End If
                    On Error GoTo 0
                Else
                    Debug.Print "    B1 值清理後與原名相同或為空，無需重新命名。"
                End If

                ' 2. 將此 B1 值標記為已處理 (表示有目標表了)
                processedSheetNames.Add b1Value, targetSheetName ' 記錄哪個表是這個 B1 的主表

                ' 3. 查找並合併其他具有相同 B1 值的 Export 表
                For secondaryWsIndex = 1 To destWb.Worksheets.Count
                    If secondaryWsIndex = wsIndex Then GoTo NextInnerSheet ' 跳過目標表本身

                    On Error Resume Next
                    Set secondarySheet = Nothing
                    Set secondarySheet = destWb.Worksheets(secondaryWsIndex)
                    If Err.Number <> 0 Then
                         Err.Clear
                         Debug.Print "    跳過內部索引 " & secondaryWsIndex & " (可能已刪除)"
                         GoTo NextInnerSheet
                    End If
                    On Error GoTo 0

                    ' 檢查是否是 Export 表，且 B1 匹配，且 B1 尚未被清空
                    If LCase(secondarySheet.Name) Like "export*" And _
                       Len(Trim(CStr(secondarySheet.Range("B1").Value))) > 0 And _
                       Trim(CStr(secondarySheet.Range("B1").Value)) = b1Value Then

                        Debug.Print "      找到需要合併數據的來源: '" & secondarySheet.Name & "'"

                        ' 計算目標表的下一個空行 (基於 A 列)
                        destNextRow = targetSheet.Cells(targetSheet.Rows.Count, "A").End(xlUp).Row + 1
                        If destNextRow < 3 Then destNextRow = 3 ' 確保至少從第3行開始貼

                        ' 計算來源表的數據範圍
                        sourceLastRow = secondarySheet.Cells(secondarySheet.Rows.Count, "A").End(xlUp).Row ' 基於 A 列
                        sourceLastCol = secondarySheet.Cells(1, secondarySheet.Columns.Count).End(xlToLeft).Column ' 基於第1行
                        If sourceLastCol < 1 Then sourceLastCol = 1

                        ' 如果來源表有數據 (第3行及以後)
                        If sourceLastRow >= 3 Then
                            Dim sourceDataRange As Range
                            On Error Resume Next
                            Set sourceDataRange = secondarySheet.Range(secondarySheet.Cells(3, "A"), secondarySheet.Cells(sourceLastRow, sourceLastCol))
                            If Err.Number = 0 Then
                                Debug.Print "        準備複製範圍: " & secondarySheet.Name & "!" & sourceDataRange.Address
                                sourceDataRange.Copy Destination:=targetSheet.Cells(destNextRow, "A")
                                Application.CutCopyMode = False
                                Debug.Print "        數據已複製到 '" & targetSheetName & "' 的第 " & destNextRow & " 行"
                            Else
                                Debug.Print "        **錯誤**: 無法定義來源數據範圍 on " & secondarySheet.Name & ". Error: " & Err.Description
                                Err.Clear
                            End If
                            Set sourceDataRange = Nothing
                            On Error GoTo 0
                        Else
                            Debug.Print "        來源工作表 '" & secondarySheet.Name & "' 無數據可複製 (第3行及以後)."
                        End If

                        ' 清除來源表的 B1 作為處理標記
                        Debug.Print "        正在清除來源 '" & secondarySheet.Name & "' 的 B1 (或其合併區域)..."
                        On Error Resume Next ' 臨時錯誤處理
                        Dim srcB1 As Range
                        Set srcB1 = secondarySheet.Range("B1")
                        
                        If srcB1.MergeCells Then
                            Dim srcMergedArea As Range
                            Set srcMergedArea = srcB1.MergeArea
                            Debug.Print "          B1 是合併儲存格 (" & srcMergedArea.Address & ")，清除整個區域..."
                            srcMergedArea.ClearContents
                            If Err.Number <> 0 Then
                                 Debug.Print "          **錯誤**: 清除來源合併區域 " & srcMergedArea.Address & " 失敗. Error: " & Err.Description
                                 Err.Clear
                            End If
                            Set srcMergedArea = Nothing
                        Else
                            ' B1 is not merged, clear it directly
                            srcB1.ClearContents
                            If Err.Number <> 0 Then
                                 Debug.Print "          **錯誤**: 清除來源未合併的 B1 失敗. Error: " & Err.Description
                                 Err.Clear
                            End If
                        End If
                        On Error GoTo 0 ' 恢復錯誤處理
                        Set srcB1 = Nothing
                        Debug.Print "        已嘗試清除來源 B1 或其合併區域。"

                    End If ' End check for secondary sheet criteria
NextInnerSheet:
                    Set secondarySheet = Nothing
                Next secondaryWsIndex ' 下一個可能的 secondary sheet

                ' 4. 清除目標表 (已改名) 的 B1 作為處理標記
                ' Debug.Print "    正在清除目標 '" & targetSheetName & "' 的 B1 (或其合併區域)..."
                ' On Error Resume Next ' 臨時錯誤處理
                ' Dim tgtB1 As Range
                ' Set tgtB1 = targetSheet.Range("B1") ' 注意 targetSheet 可能在改名後引用
                
                ' If tgtB1.MergeCells Then
                    ' Dim tgtMergedArea As Range
                    ' Set tgtMergedArea = tgtB1.MergeArea
                    ' Debug.Print "      B1 是合併儲存格 (" & tgtMergedArea.Address & ")，清除整個區域..."
                    ' tgtMergedArea.ClearContents
                    ' If Err.Number <> 0 Then
                         ' Debug.Print "      **錯誤**: 清除目標合併區域 " & tgtMergedArea.Address & " 失敗. Error: " & Err.Description
                         ' Err.Clear
                    ' End If
                    ' Set tgtMergedArea = Nothing
                ' Else
                    ' B1 is not merged, clear it directly
                    ' tgtB1.ClearContents
                    ' If Err.Number <> 0 Then
                         ' Debug.Print "      **錯誤**: 清除目標未合併的 B1 失敗. Error: " & Err.Description
                         ' Err.Clear
                    ' End If
                ' End If
                ' On Error GoTo 0 ' 恢復錯誤處理
                ' Set tgtB1 = Nothing
                Debug.Print "    已嘗試清除目標 B1 或其合併區域。"

            End If ' End check if b1Value was already processed
        End If ' End check if it's an export sheet with content in B1
NextOuterSheet:
        Set targetSheet = Nothing ' 釋放當前目標
    Next wsIndex ' 下一個可能作為目標的工作表

    Debug.Print "----- 8b 處理完成 -----"

    ' --- 在清理和顯示最終訊息之前，呼叫刪除程序 ---
    Debug.Print "----- 開始執行 DeleteExportSheets -----"
    Call DeleteExportSheets ' 或者直接寫 DeleteExportSheets 也可以
    Debug.Print "----- DeleteExportSheets 執行完畢 -----"


CleanUp:
    ' --- 9. 收尾 ---
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True ' DisplayAlerts 會在 DeleteExportSheets 內部被設為 True
    Application.Calculation = xlCalculationAutomatic

    ' Release object variables
    Set sourceFile = Nothing
    Set sourceFolder = Nothing
    Set fso = Nothing
    Set destWb = Nothing ' Keep workbook open
    Set sourceWb = Nothing
    Set sourceSheet = Nothing
    Set ws = Nothing
    Set lookupSheet = Nothing
    Set exportSheet = Nothing

    ' Release new object variables
    'Set groupValueDict = Nothing
    'Set groupSheet = Nothing
    'Set sourceDataSheet = Nothing


    ' --- 10. 完成提示 ---
    ' 注意：最終的提示訊息可能需要調整，因為 DeleteExportSheets 自身也會彈出提示
    ' 可以考慮將 DeleteExportSheets 的 MsgBox 註解掉，只保留主程序的 MsgBox
    If fileCount > 0 Then
         MsgBox "合併與處理完成！共合併了 " & fileCount & " 個檔案，並已嘗試相關後續操作。", vbInformation
         ' 原來的 MsgBox 可能會因為 DeleteExportSheets 的 MsgBox 而顯得重複
    End If

End Sub

' 清理工作表名稱的輔助函數 (保持不變)
Private Function CleanSheetName(ByVal sheetName As String) As String
    Dim invalidChars As String
    Dim i As Integer
    Dim tempChar As String

    sheetName = Trim(sheetName)
    If Len(sheetName) = 0 Then
        CleanSheetName = "_"
        Exit Function
    End If

    invalidChars = "\/?*[]:"

    For i = 1 To Len(invalidChars)
        tempChar = Mid(invalidChars, i, 1)
        sheetName = Replace(sheetName, tempChar, "_")
    Next i

    If Len(sheetName) > 31 Then
        sheetName = Left(sheetName, 31)
    End If

    If Left(sheetName, 1) = "'" Then sheetName = "_" & Mid(sheetName, 2)
    If Right(sheetName, 1) = "'" Then sheetName = Mid(sheetName, 1, Len(sheetName) - 1) & "_"

    If Len(Trim(sheetName)) = 0 Then sheetName = "_"

    CleanSheetName = sheetName
End Function


Sub DeleteExportSheets()

    ' --- 變數宣告 ---
    Dim ws As Worksheet
    Dim i As Integer
    Dim wsCount As Integer
    Dim deletedCount As Long
    Dim errorOccurred As Boolean ' 新增標記，記錄是否發生錯誤

    ' --- 執行前的基本檢查 ---
    If ActiveWorkbook Is Nothing Then
         MsgBox "沒有活動中的工作簿可操作。", vbExclamation
         Exit Sub
    End If

    ' --- 初始化 ---
    deletedCount = 0
    errorOccurred = False
    wsCount = ActiveWorkbook.Worksheets.Count

    ' --- 關鍵：確保禁用事件和提示 ---
    Application.DisplayAlerts = False
    Application.EnableEvents = False ' <--- 確保在這裡禁用事件

    ' --- 核心邏輯：從後往前遍歷並刪除 ---
    For i = wsCount To 1 Step -1
        On Error Resume Next ' <--- 在刪除操作前啟用錯誤處理
        Set ws = Nothing ' 清除之前的引用
        Set ws = ActiveWorkbook.Worksheets(i)

        If Err.Number <> 0 Then ' 獲取工作表時就出錯，跳過
            Debug.Print "獲取索引 " & i & " 的工作表時出錯: " & Err.Description
            Err.Clear
            errorOccurred = True
            GoTo NextIteration ' 使用 GoTo 跳到迴圈末尾
        End If

        ' 檢查工作表名稱是否以 "export" 開頭 (不區分大小寫)
        If LCase(Left(ws.Name, 6)) = "export" Then
            Err.Clear ' 清除之前的錯誤狀態
            ws.Delete ' 嘗試刪除工作表

            If Err.Number <> 0 Then ' <--- 檢查刪除操作是否成功
                Debug.Print "刪除工作表 '" & ws.Name & "' 時發生錯誤: " & Err.Description
                errorOccurred = True ' 標記發生了錯誤
                Err.Clear ' 清除錯誤，以便繼續循環
            Else
                ' 刪除成功
                deletedCount = deletedCount + 1
                Debug.Print "已刪除工作表: " & ws.Name
            End If
        End If
NextIteration: ' 迴圈繼續點
        Set ws = Nothing ' 釋放工作表對象
        On Error GoTo 0 ' <--- 在檢查下一個表之前，恢復正常錯誤處理模式
    Next i

    ' --- 恢復系統設置 ---
    Application.EnableEvents = True ' <--- 在 Sub 結束前恢復事件
    Application.DisplayAlerts = True

    ' --- 完成後給予使用者提示 ---
    If errorOccurred Then
        MsgBox "操作完成，共刪除了 " & deletedCount & " 個工作表。" & vbCrLf & _
               "但過程中發生了至少一個錯誤，請檢查即時運算窗口 (Ctrl+G) 獲取詳細信息。", vbExclamation
    ElseIf deletedCount > 0 Then
        MsgBox "操作完成！共刪除了 " & deletedCount & " 個以 'export' 開頭的工作表。", vbInformation
    Else
        MsgBox "在目前活動的工作簿中沒有找到以 'export' 開頭的工作表可供刪除。", vbInformation
    End If

End Sub

'========================================================================================
' Main Function: ConvertToHallNumber
' Purpose: Extracts the number from a string like "台北市召會第十二會所"
' Input:   inputText - The string containing the hall name and number.
' Output:  Variant - The extracted number (e.g., 12) or an error value if format is wrong.
' Usage (in Excel cell): =ConvertToHallNumber(A1)
' Usage (in VBA):      Dim hallNum As Variant
'                      hallNum = ConvertToHallNumber("台北市召會第三十六會所")
'========================================================================================
Public Function ConvertToHallNumber(inputText As String) As Variant
    Dim startPos As Long
    Dim endPos As Long
    Dim chineseNumStr As String
    Dim result As Variant

    ' Define the delimiters
    Const START_DELIMITER As String = "第"
    Const END_DELIMITER As String = "會所"

    ' Find the position of the delimiters
    startPos = InStr(1, inputText, START_DELIMITER, vbTextCompare) ' Use vbTextCompare for potentially different character widths
    endPos = InStr(1, inputText, END_DELIMITER, vbTextCompare)

    ' Check if both delimiters are found and in the correct order
    If startPos > 0 And endPos > startPos Then
        ' Extract the Chinese number string between the delimiters
        chineseNumStr = Mid(inputText, startPos + Len(START_DELIMITER), endPos - startPos - Len(START_DELIMITER))

        ' Try to convert the extracted Chinese numeral string to an Arabic number
        On Error Resume Next ' Enable error trapping for the conversion function
        result = ChineseToArabic(chineseNumStr)
        On Error GoTo 0 ' Disable error trapping

        ' Check if the conversion function returned an error
        If IsError(result) Then
            ConvertToHallNumber = CVErr(xlErrValue) ' Return #VALUE! if conversion failed
        Else
            ConvertToHallNumber = result ' Return the converted number
        End If
    Else
        ' Input string does not match the expected format
        ConvertToHallNumber = CVErr(xlErrNA) ' Return #N/A for format mismatch
    End If

End Function

'========================================================================================
' Helper Function: ChineseToArabic
' Purpose: Converts a Chinese numeral string (like "十二", "三十六", "一百零七") to Arabic numerals.
' Input:   chinNum - The Chinese numeral string.
' Output:  Variant - The Arabic number (Long) or an error value if conversion fails.
' Note:    This handles numbers up to 999 using 十 and 百. Simplified for the examples provided.
'          Handles both ? and 零 for zero.
'========================================================================================
Private Function ChineseToArabic(chinNum As String) As Variant
    Dim tempStr As String
    Dim resultVal As Long
    Dim posBai As Long, posShi As Long

    ' Handle empty input
    If Len(Trim(chinNum)) = 0 Then
        ChineseToArabic = CVErr(xlErrValue) ' Cannot convert empty string
        Exit Function
    End If
    
    ' --- Step 1: Replace Chinese digits with Arabic digits ---
    tempStr = chinNum
    tempStr = Replace(tempStr, "?", "0") ' Common zero
    tempStr = Replace(tempStr, "零", "0") ' Formal zero
    tempStr = Replace(tempStr, "一", "1")
    tempStr = Replace(tempStr, "二", "2")
    tempStr = Replace(tempStr, "三", "3")
    tempStr = Replace(tempStr, "四", "4")
    tempStr = Replace(tempStr, "五", "5")
    tempStr = Replace(tempStr, "六", "6")
    tempStr = Replace(tempStr, "七", "7")
    tempStr = Replace(tempStr, "八", "8")
    tempStr = Replace(tempStr, "九", "9")

    ' --- Step 2: Parse the mixed string (digits, 百, 十) ---
    On Error GoTo ConversionError ' Trap errors during Val() or logic

    resultVal = 0 ' Initialize result

    ' --- Handle Hundreds (百) ---
    posBai = InStr(1, tempStr, "百", vbTextCompare)
    If posBai > 0 Then
        Dim leftBai As String
        Dim valBai As Long
        
        leftBai = Trim(Mid(tempStr, 1, posBai - 1))
        
        If Len(leftBai) = 0 Then ' Case: "百" (implies 100) or "一百"
             valBai = 1
        ElseIf IsNumeric(leftBai) Then
             valBai = Val(leftBai)
        Else
             GoTo ConversionError ' Invalid character before 百
        End If

        resultVal = valBai * 100
        ' Process the remainder of the string after 百 for tens/units
        tempStr = Trim(Mid(tempStr, posBai + 1))
    End If ' End of handling 百

    ' --- Handle Tens (十) and Units ---
    posShi = InStr(1, tempStr, "十", vbTextCompare)
    If posShi > 0 Then
         Dim leftShi As String, rightShi As String
         Dim valShiTens As Long, valShiUnits As Long

         leftShi = Trim(Mid(tempStr, 1, posShi - 1))
         rightShi = Trim(Mid(tempStr, posShi + 1))

         ' Calculate Tens part
         If Len(leftShi) = 0 Then ' Case: "十" (implies 10) or start of string like "十二"
             valShiTens = 10
         ElseIf IsNumeric(leftShi) Then ' Case: "五十", "二十分"
             valShiTens = Val(leftShi) * 10
         Else
             GoTo ConversionError ' Invalid character before 十
         End If
         
         ' Calculate Units part
         If Len(rightShi) = 0 Then ' Case: "十" or "五十" (no units)
             valShiUnits = 0
         ElseIf IsNumeric(rightShi) Then ' Case: "十二", "五十五"
              valShiUnits = Val(rightShi)
         Else
             GoTo ConversionError ' Invalid character after 十
         End If
         
         resultVal = resultVal + valShiTens + valShiUnits ' Add tens and units to hundreds (if any)
         
    Else ' No "十" found in the remaining string (or original string)
        ' Remaining string should be just digits (units or zero)
        If Len(tempStr) > 0 Then
            If IsNumeric(tempStr) Then
                 resultVal = resultVal + Val(tempStr)
            Else
                 GoTo ConversionError ' Invalid characters remain
            End If
        End If ' If tempStr is empty here (e.g. after "一百"), resultVal is already correct
    End If ' End of handling 十

    ' Return the final calculated value
    ChineseToArabic = resultVal
    Exit Function

ConversionError:
    ' If any error occurred during parsing (e.g., non-numeric parts)
    ChineseToArabic = CVErr(xlErrValue) ' Return #VALUE!

End Function

'========================================================================================
' Example Subroutine to demonstrate usage
'========================================================================================
Sub ProcessHallData()
    Dim hallName As String
    Dim hallNumber As Variant

    hallName = "台北市召會第六十會所"
    hallNumber = ConvertToHallNumber(hallName)
    MsgBox "Convert Result: " & hallName & vbCrLf & CStr(hallNumber)

    If IsError(hallNumber) Then
        MsgBox "Could not convert hall name: " & hallName & vbCrLf & "Error: " & CStr(hallNumber)
    Else
        Range("B1").Value = hallNumber ' Put the result in cell B1
    End If
End Sub


