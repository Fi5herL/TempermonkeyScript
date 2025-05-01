Option Explicit ' 強制宣告所有變數

Sub Button2_Click()
    Call MergeSpecificExcelFiles_And_Process_V5 ' 版本號更新
End Sub

Sub MergeSpecificExcelFiles_And_Process_V5() ' 版本號更新

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
            Set sourceWb = Nothing ' Reset before opening
            Set sourceWb = Workbooks.Open(sourceFile.Path, ReadOnly:=False, UpdateLinks:=0)

            If Err.Number = 0 Then
                If sourceWb.Sheets.Count > 0 Then
                    Set sourceSheet = sourceWb.Sheets(1)
                    newSheetName = fso.GetBaseName(sourceFile.Name)
                    newSheetName = CleanSheetName(newSheetName) ' Clean name *before* potential duplication check

                    ' --- Handle potential duplicate sheet names during copy ---
                    Dim tempName As String
                    Dim nameCounter As Integer
                    tempName = newSheetName
                    nameCounter = 1
                    Do While SheetExists(destWb, tempName)
                        tempName = newSheetName & " (" & nameCounter & ")"
                        nameCounter = nameCounter + 1
                        If Len(tempName) > 31 Then ' Prevent exceeding sheet name limit
                            tempName = Left(newSheetName, 31 - Len(CStr(nameCounter)) - 3) & " (" & nameCounter & ")"
                        End If
                    Loop
                    newSheetName = tempName
                    ' --- End duplicate handling ---

                    sourceSheet.Copy After:=destWb.Sheets(destWb.Sheets.Count)
                    On Error Resume Next ' Handle error during renaming
                    destWb.Sheets(destWb.Sheets.Count).Name = newSheetName
                    If Err.Number <> 0 Then
                        Debug.Print "Error renaming sheet to '" & newSheetName & "'. Sheet might already exist or name is invalid. Error: " & Err.Description
                        ' Attempt to delete the incorrectly named sheet if renaming failed
                        On Error Resume Next ' Ignore error if delete fails
                        destWb.Sheets(destWb.Sheets.Count).Delete
                        On Error GoTo 0
                        Err.Clear
                    Else
                         fileCount = fileCount + 1
                    End If
                    On Error GoTo 0 ' Restore default error handling

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

    ' --- 8a. 後處理：查找資料並為 Export 表格添加"元數據"和"會所名稱" ---
    '     *** 這一步現在在複製後、合併前執行 ***
    Dim ws As Worksheet
    Dim lookupSheet As Worksheet
    Dim exportSheet As Worksheet
    Dim origFileNameCol As Variant
    Dim huosuoNameCol As Variant
    Dim exportTimeCol As Variant ' 新增：匯出時間欄位
    Dim itemNameCol As Variant   ' 新增：項目名稱欄位
    Dim statusNameCol As Variant ' 新增：狀態名稱欄位
    Dim foundRow As Variant
    Dim huosuoName As String     ' 會所名稱(文字)
    Dim numericHallNumber As Variant ' 會所名稱(數字)
    Dim exportTimeVal As Variant
    Dim itemNameVal As String
    Dim statusNameVal As String
    Dim lastRow As Long          ' 用於填充會所名稱列 (F列)
    Dim foundLookupSheet As Boolean
    Dim missingHeaders As Boolean

    foundLookupSheet = False
    missingHeaders = False
    Set lookupSheet = Nothing

    ' --- 8a.1: 找到有效的 "聚會資料_" 工作表並確認所需標頭 ---
    Debug.Print "----- 開始查找 Lookup Sheet -----"
    For Each ws In destWb.Worksheets
        If LCase(ws.Name) Like "聚會資料_*" Then
            Debug.Print "檢查工作表: " & ws.Name
            missingHeaders = False ' Reset for each potential lookup sheet
            On Error Resume Next ' Prevent error if Match fails
            origFileNameCol = Application.Match("原始檔案名稱", ws.Rows(1), 0)
            huosuoNameCol = Application.Match("會所名稱", ws.Rows(1), 0)
            exportTimeCol = Application.Match("匯出時間", ws.Rows(1), 0) ' 查找新標頭
            itemNameCol = Application.Match("項目名稱", ws.Rows(1), 0)   ' 查找新標頭
            statusNameCol = Application.Match("狀態名稱", ws.Rows(1), 0) ' 查找新標頭
            On Error GoTo 0 ' Restore error handling

            ' 檢查所有必要的標頭是否都找到
            If IsError(origFileNameCol) Then Debug.Print "  未找到 '原始檔案名稱' 標頭"; missingHeaders = True
            If IsError(huosuoNameCol) Then Debug.Print "  未找到 '會所名稱' 標頭"; missingHeaders = True
            If IsError(exportTimeCol) Then Debug.Print "  未找到 '匯出時間' 標頭"; missingHeaders = True
            If IsError(itemNameCol) Then Debug.Print "  未找到 '項目名稱' 標頭"; missingHeaders = True
            If IsError(statusNameCol) Then Debug.Print "  未找到 '狀態名稱' 標頭"; missingHeaders = True

            If Not missingHeaders Then
                Set lookupSheet = ws
                foundLookupSheet = True
                Debug.Print "  找到包含所有必要標頭的 Lookup Sheet: " & lookupSheet.Name
                Exit For ' 找到第一個符合條件的就跳出
            Else
                 Debug.Print "  工作表 '" & ws.Name & "' 缺少部分或全部必要標頭，繼續尋找..."
                 ' Reset variables to ensure a fully valid sheet is used later
                 origFileNameCol = Empty: huosuoNameCol = Empty: exportTimeCol = Empty: itemNameCol = Empty: statusNameCol = Empty
            End If
        End If
    Next ws
    Debug.Print "----- Lookup Sheet 查找結束 -----"

    ' --- 8a.2: 遍歷 Export 工作表，插入欄位並填入數據 ---
    If foundLookupSheet Then
        Debug.Print "----- 開始處理 Export Sheets (插入元數據和會所列) -----"
        For Each exportSheet In destWb.Worksheets
            If LCase(exportSheet.Name) Like "export*" Then
                Debug.Print "處理 Export Sheet: " & exportSheet.Name

                ' --- 查找對應行 ---
                On Error Resume Next
                foundRow = Application.Match(exportSheet.Name, lookupSheet.Columns(origFileNameCol), 0)
                On Error GoTo 0

                If Not IsError(foundRow) Then
                    ' --- 找到匹配行 ---
                    Debug.Print "  找到匹配行: " & foundRow & " in " & lookupSheet.Name
                    ' --- 提取元數據 ---
                    exportTimeVal = lookupSheet.Cells(foundRow, exportTimeCol).Value
                    huosuoName = CStr(lookupSheet.Cells(foundRow, huosuoNameCol).Value)
                    itemNameVal = CStr(lookupSheet.Cells(foundRow, itemNameCol).Value)
                    statusNameVal = CStr(lookupSheet.Cells(foundRow, statusNameCol).Value)

                    ' --- 轉換會所名稱為數字 ---
                    numericHallNumber = ConvertToHallNumber(huosuoName)
                    If IsError(numericHallNumber) Then
                        Debug.Print "    **警告**: 無法轉換會所名稱 '" & huosuoName & "' 為數字. Error: " & CStr(numericHallNumber)
                        numericHallNumber = "#N/A" ' 使用錯誤標識
                    Else
                        Debug.Print "    會所名稱 '" & huosuoName & "' 轉換為: " & numericHallNumber
                    End If

                    ' --- 插入 A:F 列 (共6列) ---
                    exportSheet.Columns("A:F").Insert Shift:=xlToRight
                    Debug.Print "    已插入 A:F 欄"

                    ' --- 填寫 第 2 列 的標題 (A2:E2 for metadata, F2 for hall name) ---
                    exportSheet.Range("A2:E2").Value = Array("匯出時間", "會所中文名稱", "項目名稱", "狀態名稱", "原始檔案名稱")
                    exportSheet.Range("F2").Value = "會所名稱"
                    Debug.Print "    已填入第 2 行標題 (A2:F2)"

                    ' --- 填寫 第 3 列 的元數據值 (A3:E3) ---
                    exportSheet.Range("A3").Value = exportTimeVal ' 直接賦值保留格式
                    exportSheet.Range("B3").Value = huosuoName ' 使用中文會所名稱
                    exportSheet.Range("C3").Value = itemNameVal
                    exportSheet.Range("D3").Value = statusNameVal
                    exportSheet.Range("E3").Value = exportSheet.Name ' 分頁名即原始檔名
                    ' 嘗試格式化日期
                    On Error Resume Next
                    exportSheet.Range("A3").NumberFormat = "yyyy/m/d h:mm"
                    On Error GoTo 0
                    Debug.Print "    已填入第 3 行元數據值 (A3:E3)"

                    ' --- 填充 第 F 列 的數字會所名稱 (從 F3 開始) ---
                    ' 找到原始數據的最後一行 (現在原始數據從 G 欄開始)
                    ' 數據現在從第 3 行開始
                    If exportSheet.Cells(exportSheet.Rows.Count, "G").End(xlUp).Row < 3 Then
                        ' 如果原始數據區域 (G欄起) 在第3行之前就結束了 (即只有標頭或完全空白)
                        lastRow = 2 ' 設置為 2 使得後續 >= 3 的判斷為 False
                        Debug.Print "    原始數據區域 (G欄起) 在第 3 行或之前結束，無需填充 F 列數據."
                    Else
                        lastRow = exportSheet.Cells(exportSheet.Rows.Count, "G").End(xlUp).Row
                        Debug.Print "    原始數據區域 (G欄起) 最後一列為: " & lastRow & " (數據從第 3 行開始)"
                    End If

                    ' 從 F3 開始向下填充
                    If lastRow >= 3 Then
                        exportSheet.Range("F3:F" & lastRow).Value = numericHallNumber ' 使用之前轉換好的數字
                        exportSheet.Range("A3:A" & lastRow).Value = exportTimeVal ' 使用之前轉換好的數字
                        exportSheet.Range("B3:B" & lastRow).Value = huosuoName ' 使用之前轉換好的數字
                        exportSheet.Range("C3:C" & lastRow).Value = itemNameVal ' 使用之前轉換好的數字
                        exportSheet.Range("D3:D" & lastRow).Value = statusNameVal ' 使用之前轉換好的數字
                        exportSheet.Range("E3:E" & lastRow).Value = exportSheet.Name ' 使用之前轉換好的數字
                        Debug.Print "    已將數字會所名稱填入 F3:F" & lastRow
                    End If

                Else
                    ' --- 未找到匹配行 ---
                    Debug.Print "  **警告**: 在 Lookup Sheet 中未找到 '" & exportSheet.Name & "' 的匹配項."
                    ' --- 仍然插入 A:F 列並填寫標題和標記 ---
                    exportSheet.Columns("A:F").Insert Shift:=xlToRight
                    exportSheet.Range("A2:E2").Value = Array("匯出時間", "會所名稱", "項目名稱", "狀態名稱", "原始檔案名稱")
                    exportSheet.Range("F2").Value = "會所名稱"
                    exportSheet.Range("A3:E3").Value = Array("N/A", "未找到", "N/A", "N/A", exportSheet.Name) ' 標示未找到
                    exportSheet.Range("F3").Value = "未找到" ' 在 F3 標示
                    Debug.Print "    已插入 A:F 欄並標示未找到數據 (A3:F3)"
                End If
            End If ' End If LCase(exportSheet.Name) Like "export*"
        Next exportSheet
        Debug.Print "----- Export Sheets 處理完成 (插入元數據和會所列) -----"
    Else
        MsgBox "警告: 未找到有效的 '聚會資料_' 查找表（包含所有必需的標頭）。" & vbCrLf & _
               "無法為 Export 分頁添加元數據和會所名稱列。", vbExclamation
        ' 即使沒有 lookup sheet，仍然繼續後面的合併和刪除步驟
    End If

    ' --- 8b. 修改後邏輯：重新命名第一個，合併其餘相同 G2 (原 B1) 的 ---
    '     *** 這部分現在處理已經插入了 A:F 列的表 ***
    Dim targetSheet As Worksheet
    Dim secondarySheet As Worksheet
    Dim g2Value As String             ' 使用 G2 (原 B1, 現在是分組依據的標頭行)
    Dim groupValue As String          ' <--- 使用新變數名
    Dim targetSheetName As String
    Dim processedSheetNames As Object
    Dim sourceLastRow As Long
    Dim sourceLastCol As Long
    Dim destNextRow As Long
    Dim wsIndex As Integer
    Dim secondaryWsIndex As Integer
    Dim dataStartCol As String
    Dim groupHeaderRow As Long
    Dim dataStartRow As Long

    dataStartCol = "G"          ' 原始數據現在從 G 開始
    groupHeaderRow = 1          ' 用於分組的值所在的行 (原 G1
    dataStartRow = 3            ' 實際數據開始的行 (原 Row 2, 新Row 3

    Set processedSheetNames = CreateObject("Scripting.Dictionary")

    Debug.Print "----- 開始 8b：重新命名與合併 (基於 " & dataStartCol & groupHeaderRow & ") -----"

    For wsIndex = 1 To destWb.Worksheets.Count
        On Error Resume Next
        Set targetSheet = Nothing
        Set targetSheet = destWb.Worksheets(wsIndex)
        If Err.Number <> 0 Then
            Err.Clear
            Debug.Print "  跳過索引 " & wsIndex & " (可能已刪除)"
            GoTo NextOuterSheet8b ' 跳到下一個外層循環
        End If
        On Error GoTo 0

        ' 檢查是否是 Export 表 且 G2 (原 B1) 尚未處理 (G2 有內容)
        ' LCase 檢查仍然基於原始名稱，因為重命名發生在此步驟內部
        ' 注意: 必須確保 targetSheet 在步驟 8a 中被處理過 (即 A:F 列已插入)
        ' 為簡化，假設所有 export* 表都被 8a 處理了 (即使是 Not Found 情況)
        If LCase(Left(targetSheet.Name, 6)) = "export" And Len(Trim(CStr(targetSheet.Range(dataStartCol & groupHeaderRow).Value))) > 0 Then

            g2Value = Trim(CStr(targetSheet.Range(dataStartCol & groupHeaderRow).Value))
            targetSheetName = targetSheet.Name ' 記錄當前名稱

            If Not processedSheetNames.Exists(g2Value) Then
                Debug.Print "  找到新的目標分組 " & dataStartCol & groupHeaderRow & ": '" & g2Value & "', 來源工作表: '" & targetSheetName & "'"

                ' --- 1. 重新命名目標工作表 ---
                Dim newNameCandidate As String
                newNameCandidate = CleanSheetName(g2Value)
                If Len(newNameCandidate) > 0 And LCase(newNameCandidate) <> LCase(targetSheetName) Then
                    On Error Resume Next
                    targetSheet.Name = newNameCandidate
                    If Err.Number <> 0 Then
                        Debug.Print "    **警告**: 無法將工作表 '" & targetSheetName & "' 重新命名為 '" & newNameCandidate & "'. Error: " & Err.Description & ". 將維持原名處理。"
                        Err.Clear
                        ' 如果重命名失敗，則此分組無法正確合併，但標記為已處理避免後續衝突
                    Else
                        targetSheetName = newNameCandidate
                        Debug.Print "    工作表已重新命名為: '" & targetSheetName & "'"
                    End If
                    On Error GoTo 0
                Else
                    Debug.Print "    " & dataStartCol & groupHeaderRow & " 值清理後與原名相同或為空，無需重新命名。"
                End If

                ' --- 2. 將此 G2 值標記為已處理 ---
                processedSheetNames.Add g2Value, targetSheetName

                ' --- 3. 查找並合併其他具有相同 G2 值的 Export 表 ---
                For secondaryWsIndex = 1 To destWb.Worksheets.Count
                    If secondaryWsIndex = wsIndex Then GoTo NextInnerSheet8b

                    On Error Resume Next
                    Set secondarySheet = Nothing
                    Set secondarySheet = destWb.Worksheets(secondaryWsIndex)
                    If Err.Number <> 0 Then
                         Err.Clear
                         Debug.Print "    跳過內部索引 " & secondaryWsIndex & " (可能已刪除)"
                         GoTo NextInnerSheet8b
                    End If
                    On Error GoTo 0

                    ' 檢查是否是 Export 表，且 G2 匹配，且 G2 尚未被清空
                    If LCase(Left(secondarySheet.Name, 6)) = "export" And _
                       Len(Trim(CStr(secondarySheet.Range(dataStartCol & groupHeaderRow).Value))) > 0 And _
                       Trim(CStr(secondarySheet.Range(dataStartCol & groupHeaderRow).Value)) = g2Value Then

                        Debug.Print "      找到需要合併數據的來源: '" & secondarySheet.Name & "'"

                        ' --- 計算目標表的下一個空行 (基於 G 列，數據從第 3 行開始) ---
                        destNextRow = targetSheet.Cells(targetSheet.Rows.Count, dataStartCol).End(xlUp).Row + 1
                        If destNextRow < dataStartRow Then destNextRow = dataStartRow ' 確保至少從第 3 行開始貼
                        Debug.Print "        目標表 '" & targetSheetName & "' 下一空行 (基於 " & dataStartCol & "): " & destNextRow

                        ' --- 計算來源表的數據範圍 (從第 3 行開始) ---
                        sourceLastRow = secondarySheet.Cells(secondarySheet.Rows.Count, dataStartCol).End(xlUp).Row
                        ' 找到數據行的最後一列 (基於第 3 行)
                        sourceLastCol = 0
                        If sourceLastRow >= dataStartRow Then
                             On Error Resume Next ' Handle potentially empty row 3
                             sourceLastCol = secondarySheet.Cells(dataStartRow, secondarySheet.Columns.Count).End(xlToLeft).Column
                             If Err.Number <> 0 Or sourceLastCol < Columns(dataStartCol).Column Then
                                 sourceLastCol = Columns(dataStartCol).Column ' Default to G if error or empty
                                 Err.Clear
                             End If
                             On Error GoTo 0
                        Else
                             sourceLastCol = Columns(dataStartCol).Column ' No data rows, just use G
                        End If


                        ' --- 如果來源表有數據 (第 3 行及以後) ---
                        If sourceLastRow >= dataStartRow Then
                            Dim sourceDataRange As Range
                            'Dim sourceHallColRange As Range
                            'Dim targetHallColRange As Range

                            ' --- *** 關鍵修正：定義來源數據範圍從 A 欄開始 *** ---
                            ' --- 複製原始數據 (G 列到最後一列) ---
                            On Error Resume Next
                            Set sourceDataRange = secondarySheet.Range(secondarySheet.Cells(dataStartRow, 1), secondarySheet.Cells(sourceLastRow, sourceLastCol))
                            If Err.Number <> 0 Or sourceDataRange Is Nothing Then
                                Debug.Print "        **錯誤**: 無法定義來源數據範圍 on " & secondarySheet.Name & ". Error: " & Err.Description & ". Or Range is Nothing."
                                Err.Clear
                            Else
                            Debug.Print "        準備複製數據範圍: " & secondarySheet.Name & "!" & sourceDataRange.Address
                                sourceDataRange.Copy Destination:=targetSheet.Cells(destNextRow, 1)
                                Application.CutCopyMode = False
                                Debug.Print "        數據已複製到 '" & targetSheetName & "' 的第 " & destNextRow & " 行, 從 A 欄開始"

                                ' --- 不再需要單獨複製 F 欄 ---
                               ' Set sourceHallColRange = secondarySheet.Range("F" & dataStartRow & ":F" & sourceLastRow)
                               ''Set targetHallColRange = targetSheet.Range("F" & destNextRow & ":F" & (destNextRow + sourceHallColRange.Rows.Count - 1))
                                'If sourceHallColRange.Rows.Count = targetHallColRange.Rows.Count Then
                                '    sourceHallColRange.Copy Destination:=targetHallColRange
                                '    Application.CutCopyMode = False
                                '    Debug.Print "        會所名稱 (F欄) 已從來源複製到目標行 " & destNextRow
                                'Else
                                '     Debug.Print "        **警告**: F欄複製時行數不匹配. Source: " & sourceHallColRange.Address & ", Target: " & targetHallColRange.Address
                                'End If
                                'Set sourceHallColRange = Nothing
                                'Set targetHallColRange = Nothing
                            'Else
                            '    Debug.Print "        **錯誤**: 無法定義來源數據範圍 on " & secondarySheet.Name & ". Error: " & Err.Description & ". Or Range is Nothing."
                            '    Err.Clear
                            End If
                            Set sourceDataRange = Nothing
                            On Error GoTo 0
                        Else
                            Debug.Print "        來源工作表 '" & secondarySheet.Name & "' 無數據可複製 (第 " & dataStartRow & " 行及以後)."
                        End If

                        ' --- 清除來源表的 G2 作為處理標記 ---
                        Debug.Print "        正在清除來源 '" & secondarySheet.Name & "' 的 " & dataStartCol & groupHeaderRow & "..."
                        On Error Resume Next ' 處理合併單元格等情況
                        Dim srcG2 As Range
                        Set srcG2 = secondarySheet.Range(dataStartCol & groupHeaderRow)
                        If Not srcG2 Is Nothing Then
                            If srcG2.MergeCells Then
                                srcG2.MergeArea.ClearContents
                                Debug.Print "          已清除合併區域 " & srcG2.MergeArea.Address
                            Else
                                srcG2.ClearContents
                                Debug.Print "          已清除單元格 " & srcG2.Address
                            End If
                            If Err.Number <> 0 Then
                                Debug.Print "          **錯誤**: 清除來源 " & dataStartCol & groupHeaderRow & " 失敗. Error: " & Err.Description
                                Err.Clear
                            End If
                        End If
                        On Error GoTo 0
                        Set srcG2 = Nothing

                    End If
NextInnerSheet8b:
                    Set secondarySheet = Nothing
                Next secondaryWsIndex

                ' --- 4. 清除目標表 (已改名) 的 G2 作為處理標記 (可選，如果需要防止它被再次選為目標) ---
                ' Debug.Print "    正在清除目標 '" & targetSheetName & "' 的 " & dataStartCol & groupHeaderRow & "..."
                ' On Error Resume Next
                ' targetSheet.Range(dataStartCol & groupHeaderRow).ClearContents
                ' If Err.Number <> 0 Then Debug.Print "      **錯誤**: 清除目標 " & dataStartCol & groupHeaderRow & " 失敗: " & Err.Description; Err.Clear
                ' On Error GoTo 0
                ' Debug.Print "    已嘗試清除目標 " & dataStartCol & groupHeaderRow

            End If ' End If Not processedSheetNames.Exists(g2Value)
        End If ' End If LCase(Left(targetSheet.Name, 6)) = "export" ...
NextOuterSheet8b:
        Set targetSheet = Nothing
    Next wsIndex

    Debug.Print "----- 8b 處理完成 -----"

    ' --- 9. 刪除處理過的 Export 表 ---
    '    DeleteExportSheets 會刪除所有名字仍以 "export" 開頭的表
    '    (即未被成功重命名或未被合併的表)
    Debug.Print "----- 開始執行 DeleteExportSheets -----"
    Call DeleteExportSheets
    Debug.Print "----- DeleteExportSheets 執行完畢 -----"

CleanUp:
    ' --- 10. 收尾 ---
    On Error Resume Next ' Prevent finalization errors from stopping msgbox
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True ' 確保恢復
    On Error GoTo 0

    ' Release object variables
    Set sourceFile = Nothing
    Set sourceFolder = Nothing
    Set fso = Nothing
    ' Keep destWb open
    Set sourceWb = Nothing
    Set sourceSheet = Nothing
    Set ws = Nothing
    Set lookupSheet = Nothing
    Set exportSheet = Nothing
    Set targetSheet = Nothing
    Set secondarySheet = Nothing
    Set processedSheetNames = Nothing

    ' --- 11. 完成提示 ---
    If Not destWb Is Nothing Then
        If fileCount > 0 Then
            MsgBox "合併與處理完成！共合併了 " & fileCount & " 個檔案。" & vbCrLf & _
                   "已為 Export 分頁添加元數據(第2行標題,第3行值)和會所名稱列(F列)，並嘗試合併與清理。", vbInformation
        ' Else: Message already shown if no files found or workbook closed early
        End If
    End If

End Sub

' === 輔助函數區域 ===

' 檢查工作表是否存在
Private Function SheetExists(wb As Workbook, sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = (Not ws Is Nothing)
End Function

' 清理工作表名稱
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

    If Left(sheetName, 1) = "'" Then sheetName = "_" & Mid(sheetName, 2)
    If Right(sheetName, 1) = "'" Then sheetName = Left(sheetName, Len(sheetName) - 1) & "_"

    If Len(sheetName) > 31 Then
        sheetName = Left(sheetName, 31)
    End If

    If Len(Trim(sheetName)) = 0 Then sheetName = "_"

    CleanSheetName = sheetName
End Function

' 刪除 Export Sheets
Sub DeleteExportSheets()
    Dim ws As Worksheet
    Dim i As Integer
    Dim wsCount As Integer
    Dim deletedCount As Long
    Dim errorOccurred As Boolean
    Dim currentWb As Workbook

    On Error Resume Next ' Handle case where ActiveWorkbook is Nothing
    Set currentWb = ActiveWorkbook
    On Error GoTo 0
    If currentWb Is Nothing Then
         Debug.Print "DeleteExportSheets: 沒有活動中的工作簿可操作。"
         Exit Sub
    End If

    deletedCount = 0
    errorOccurred = False
    wsCount = currentWb.Worksheets.Count

    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Debug.Print "--- Starting DeleteExportSheets ---"
    For i = wsCount To 1 Step -1
        On Error Resume Next
        Set ws = Nothing
        Set ws = currentWb.Worksheets(i)
        If Err.Number <> 0 Then
            Debug.Print "  Error getting sheet at index " & i & ": " & Err.Description
            Err.Clear
            errorOccurred = True
            GoTo NextDeleteIteration
        End If
        On Error GoTo 0

        If LCase(Left(ws.Name, 6)) = "export" Then
            Debug.Print "  Attempting to delete: '" & ws.Name & "'"
            On Error Resume Next
            ws.Delete
            If Err.Number <> 0 Then
                Debug.Print "    **ERROR** deleting sheet '" & ws.Name & "': " & Err.Description
                errorOccurred = True
                Err.Clear
            Else
                deletedCount = deletedCount + 1
                Debug.Print "    Successfully deleted."
            End If
            On Error GoTo 0
        End If
NextDeleteIteration:
        Set ws = Nothing
    Next i
    Debug.Print "--- Finished DeleteExportSheets ---"

    Application.EnableEvents = True
    Application.DisplayAlerts = True

    ' Use Debug.Print instead of MsgBox to avoid double prompts
    If errorOccurred Then
        Debug.Print "DeleteExportSheets: 操作完成，共刪除了 " & deletedCount & " 個工作表。但過程中發生錯誤 (詳見上方)。"
    ElseIf deletedCount > 0 Then
         Debug.Print "DeleteExportSheets: 操作完成！共刪除了 " & deletedCount & " 個以 'export' 開頭的工作表。"
    Else
         Debug.Print "DeleteExportSheets: 沒有找到以 'export' 開頭的工作表可供刪除。"
    End If
End Sub


' 會所名稱轉數字
Public Function ConvertToHallNumber(inputText As String) As Variant
    Dim startPos As Long
    Dim endPos As Long
    Dim chineseNumStr As String
    Dim result As Variant
    Const START_DELIMITER As String = "第"
    Const END_DELIMITER As String = "會所"

    startPos = InStr(1, inputText, START_DELIMITER, vbTextCompare)
    endPos = InStr(1, inputText, END_DELIMITER, vbTextCompare)

    If startPos > 0 And endPos > startPos Then
        chineseNumStr = Mid(inputText, startPos + Len(START_DELIMITER), endPos - startPos - Len(START_DELIMITER))
        On Error Resume Next
        result = ChineseToArabic(chineseNumStr)
        On Error GoTo 0
        If IsError(result) Then
            ConvertToHallNumber = CVErr(xlErrValue) ' #VALUE!
        Else
            ConvertToHallNumber = result
        End If
    Else
        If IsNumeric(inputText) Then
            On Error Resume Next ' Handle potential overflow if input is large number string
            result = CLng(inputText)
            If Err.Number <> 0 Then
                 ConvertToHallNumber = CVErr(xlErrValue) ' Conversion to Long failed
                 Err.Clear
            Else
                 ConvertToHallNumber = result
            End If
            On Error GoTo 0
        Else
             ConvertToHallNumber = CVErr(xlErrNA) ' #N/A
        End If
    End If
End Function

' 中文數字轉阿拉伯數字
Private Function ChineseToArabic(chinNum As String) As Variant
    Dim tempStr As String
    Dim resultVal As Long
    Dim posBai As Long, posShi As Long

    tempStr = Trim(chinNum)
    If Len(tempStr) = 0 Then
        ChineseToArabic = CVErr(xlErrValue)
        Exit Function
    End If

    tempStr = Replace(tempStr, "〇", "0")
    tempStr = Replace(tempStr, "零", "0")
    tempStr = Replace(tempStr, "一", "1")
    tempStr = Replace(tempStr, "二", "2")
    tempStr = Replace(tempStr, "三", "3")
    tempStr = Replace(tempStr, "四", "4")
    tempStr = Replace(tempStr, "五", "5")
    tempStr = Replace(tempStr, "六", "6")
    tempStr = Replace(tempStr, "七", "7")
    tempStr = Replace(tempStr, "八", "8")
    tempStr = Replace(tempStr, "九", "9")

    On Error GoTo ConversionError_CAToA

    resultVal = 0

    posBai = InStr(1, tempStr, "百", vbTextCompare)
    If posBai > 0 Then
        Dim leftBai As String, valBai As Long
        leftBai = Trim(Mid(tempStr, 1, posBai - 1))
        If Len(leftBai) = 0 Then valBai = 1 Else If IsNumeric(leftBai) Then valBai = Val(leftBai) Else GoTo ConversionError_CAToA
        resultVal = valBai * 100
        tempStr = Trim(Mid(tempStr, posBai + 1))
    End If

    posShi = InStr(1, tempStr, "十", vbTextCompare)
    If posShi > 0 Then
         Dim leftShi As String, rightShi As String, valShiTens As Long, valShiUnits As Long
         leftShi = Trim(Mid(tempStr, 1, posShi - 1))
         rightShi = Trim(Mid(tempStr, posShi + 1))
         If Len(leftShi) = 0 Then valShiTens = 10 Else If IsNumeric(leftShi) Then valShiTens = Val(leftShi) * 10 Else GoTo ConversionError_CAToA
         If Len(rightShi) = 0 Then valShiUnits = 0 Else If IsNumeric(rightShi) Then valShiUnits = Val(rightShi) Else GoTo ConversionError_CAToA
         resultVal = resultVal + valShiTens + valShiUnits
    Else
        If Len(tempStr) > 0 Then
            If IsNumeric(tempStr) Then resultVal = resultVal + Val(tempStr) Else GoTo ConversionError_CAToA
        End If
    End If

    ChineseToArabic = resultVal
    Exit Function

ConversionError_CAToA:
    ChineseToArabic = CVErr(xlErrValue)
End Function

' 示範用的子程序 (保持不變)
Sub ProcessHallData()
    Dim hallName As String
    Dim hallNumber As Variant

    hallName = "台北市召會第六十會所"
    hallNumber = ConvertToHallNumber(hallName)
    If IsError(hallNumber) Then MsgBox "無法轉換 '" & hallName & "'. 錯誤碼: " & CStr(hallNumber) Else MsgBox "'" & hallName & "' 轉換結果: " & hallNumber

    hallName = "測試123"
    hallNumber = ConvertToHallNumber(hallName)
    If IsError(hallNumber) Then MsgBox "無法轉換 '" & hallName & "'. 錯誤碼: " & CStr(hallNumber) Else MsgBox "'" & hallName & "' 轉換結果: " & hallNumber

    hallName = "15"
    hallNumber = ConvertToHallNumber(hallName)
    If IsError(hallNumber) Then MsgBox "無法轉換 '" & hallName & "'. 錯誤碼: " & CStr(hallNumber) Else MsgBox "'" & hallName & "' 轉換結果: " & hallNumber
End Sub
