Option Explicit

Sub Button4_Click()
    Call RenameExportFilesSequentially
End Sub

Sub RenameExportFilesSequentially()

    Dim fso As Object ' FileSystemObject
    Dim sourceFolder As Object ' Folder Object
    Dim file As Object ' File Object
    Dim folderPath As String
    Dim baseName As String
    Dim extName As String
    Dim numericPart As String
    Dim i As Long, j As Long ' Loop counters
    Dim filesToSort As Object ' Use a Dictionary to store NumericValue -> FullPath
    Dim sortedKeys As Variant ' Array to hold sorted numeric keys
    Dim key As Variant
    Dim renameCounter As Long
    Dim originalPath As String
    Dim newBaseName As String
    Dim newFullName As String
    Dim targetPath As String
    Dim filesProcessed As Long
    Dim filesMatched As Long
    Dim filesRenamed As Long
    Dim filesSkipped As Long ' For skipping rename due to errors or existing target
    Dim errorsEncountered As Long
    Dim logMessage As String

    ' --- 1. 建立 FileSystemObject & 設定引用 ---
    On Error Resume Next ' Temporary ignore errors for FSO creation check
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Then
        MsgBox "無法建立 FileSystemObject 物件。" & vbCrLf & _
               "請確認 'Microsoft Scripting Runtime' 引用已啟用 (工具 > 設定引用項目)。", vbCritical, "錯誤"
        Exit Sub
    End If
    On Error GoTo ErrorHandler ' Restore proper error handling

    ' --- 2. 選擇資料夾 ---
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇包含 'export - *.xls' 檔案的資料夾"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            MsgBox "未選擇資料夾，操作已取消。", vbExclamation, "取消"
            GoTo CleanUp
        End If
        folderPath = .SelectedItems(1)
    End With

    If Not fso.FolderExists(folderPath) Then
        MsgBox "選擇的資料夾不存在：" & vbCrLf & folderPath, vbCritical, "錯誤"
        GoTo CleanUp
    End If

    Set sourceFolder = fso.GetFolder(folderPath)

    ' --- 3. 初始化 ---
    Set filesToSort = CreateObject("Scripting.Dictionary")
    filesToSort.CompareMode = vbTextCompare ' Case-insensitive keys (though keys are numeric here)
    renameCounter = 101 ' 起始編號
    filesProcessed = 0
    filesMatched = 0
    filesRenamed = 0
    filesSkipped = 0
    errorsEncountered = 0
    logMessage = "處理記錄：" & vbCrLf & String(30, "-") & vbCrLf

    ' --- 4. 第一階段：尋找符合條件的檔案並提取數字 ---
    If sourceFolder.Files.Count = 0 Then
        MsgBox "選擇的資料夾中沒有檔案。", vbInformation, "提示"
        GoTo CleanUp
    End If

    For Each file In sourceFolder.Files
        filesProcessed = filesProcessed + 1
        baseName = fso.GetBaseName(file.Name)
        extName = LCase(fso.GetExtensionName(file.Name)) ' 使用小寫比較副檔名

        ' 檢查是否符合 "export - *.xls" 格式
        If LCase(baseName) Like "export - *" And extName = "xls" Then
            filesMatched = filesMatched + 1
            numericPart = ""

            ' 從基本檔名中提取數字
            For i = 1 To Len(baseName)
                Dim char As String
                char = Mid(baseName, i, 1)
                If IsNumeric(char) Then
                    numericPart = numericPart & char
                End If
            Next i

            If Len(numericPart) > 0 Then
                ' 嘗試將數字字串轉換為 Double 以便正確排序 (處理可能很長的數字)
                Dim numericValue As Double
                On Error Resume Next ' Handle potential conversion error
                numericValue = CDbl(numericPart)
                If Err.Number <> 0 Then
                    ' 如果轉換失敗 (非常不可能，但以防萬一)，記錄錯誤並跳過此檔案
                    logMessage = logMessage & "錯誤: 無法轉換 '" & numericPart & "' 為數字 (來自檔案: " & file.Name & "). 已跳過。" & vbCrLf
                    errorsEncountered = errorsEncountered + 1
                    Err.Clear
                Else
                    ' 檢查是否有重複的數字 (不太可能，但字典鍵必須唯一)
                    If Not filesToSort.Exists(numericValue) Then
                        filesToSort.Add numericValue, file.Path ' Key: Numeric Value, Item: Full Path
                    Else
                        ' 如果數字重複，表示時間戳可能完全相同或提取有誤
                         logMessage = logMessage & "警告: 發現重複的數字標識 '" & numericPart & "' (來自檔案: " & file.Name & " 和 " & fso.GetFileName(filesToSort(numericValue)) & "). 已跳過後者。" & vbCrLf
                         filesSkipped = filesSkipped + 1
                    End If
                End If
                On Error GoTo ErrorHandler ' Restore global error handling
            Else
                logMessage = logMessage & "跳過: " & file.Name & " (未在 'export - ' 部分後找到數字)" & vbCrLf
                filesSkipped = filesSkipped + 1
            End If
        End If ' End check for "export - *.xls"
    Next file

    ' --- 5. 檢查是否有符合條件的檔案 ---
    If filesToSort.Count = 0 Then
        MsgBox "在資料夾中未找到符合 'export - *.xls' 格式且包含可提取數字的檔案。", vbInformation, "無符合檔案"
        GoTo CleanUp
    End If

    ' --- 6. 排序數字鍵 ---
    sortedKeys = filesToSort.Keys ' 取得所有數字鍵
    ' 執行氣泡排序 (Bubble Sort) 來對數字鍵進行排序
    Dim tempKey As Variant
    For i = LBound(sortedKeys) To UBound(sortedKeys) - 1
        For j = i + 1 To UBound(sortedKeys)
            If sortedKeys(i) > sortedKeys(j) Then ' 數字比較
                tempKey = sortedKeys(i)
                sortedKeys(i) = sortedKeys(j)
                sortedKeys(j) = tempKey
            End If
        Next j
    Next i

    ' --- 7. 第二階段：根據排序結果循序重新命名 ---
    logMessage = logMessage & String(30, "-") & vbCrLf & "開始重新命名 (從 " & renameCounter & " 開始):" & vbCrLf

    For i = LBound(sortedKeys) To UBound(sortedKeys)
        key = sortedKeys(i)
        originalPath = filesToSort(key) ' 根據排序後的鍵取得原始完整路徑

        ' 再次檢查原始檔案是否存在 (以防萬一在處理過程中被刪除)
        If fso.FileExists(originalPath) Then
            newBaseName = "export (" & renameCounter & ")"
            newFullName = newBaseName & ".xls" ' 保留 .xls 副檔名
            targetPath = fso.BuildPath(folderPath, newFullName)

            ' 檢查目標檔名是否已存在 (若已存在則跳過，避免覆蓋)
            If Not fso.FileExists(targetPath) Then
                ' 執行重新命名 (使用 MoveFile)
                On Error Resume Next ' 處理可能的重新命名錯誤 (例如檔案被鎖定)
                fso.MoveFile originalPath, targetPath
                If Err.Number = 0 Then
                    logMessage = logMessage & "成功: " & fso.GetFileName(originalPath) & " -> " & newFullName & vbCrLf
                    filesRenamed = filesRenamed + 1
                Else
                    logMessage = logMessage & "錯誤: 無法將 " & fso.GetFileName(originalPath) & " 重新命名為 " & newFullName & ". 原因: " & Err.Description & vbCrLf
                    errorsEncountered = errorsEncountered + 1
                    Err.Clear ' 清除錯誤狀態
                End If
                On Error GoTo ErrorHandler ' 恢復全域錯誤處理
            Else
                logMessage = logMessage & "跳過: " & fso.GetFileName(originalPath) & " (目標檔名 " & newFullName & " 已存在)" & vbCrLf
                filesSkipped = filesSkipped + 1
            End If

            renameCounter = renameCounter + 1 ' 無論成功與否，只要處理過一個排序項目，計數器就增加
        Else
             logMessage = logMessage & "跳過: 原始檔案 '" & fso.GetFileName(originalPath) & "' 在重新命名時未找到。" & vbCrLf
             filesSkipped = filesSkipped + 1
        End If
    Next i

    ' --- 8. 顯示結果 ---
    Dim summary As String
    summary = "處理完成！" & vbCrLf & vbCrLf & _
              "總共掃描檔案數: " & filesProcessed & vbCrLf & _
              "符合 'export - *.xls' 格式: " & filesMatched & vbCrLf & _
              "成功重新命名: " & filesRenamed & vbCrLf & _
              "跳過處理/重新命名: " & filesSkipped + (filesMatched - filesToSort.Count) & " (包含格式符合但無數字/數字重複/目標已存在/未找到)" & vbCrLf & _
              "處理中發生錯誤: " & errorsEncountered & vbCrLf & vbCrLf & _
              "詳細記錄:" & vbCrLf & String(30, "-") & vbCrLf & logMessage

    MsgBox summary, vbInformation, "處理結果"
    Debug.Print summary ' 將完整記錄輸出到 VBE 的 即時運算視窗 (Ctrl+G)

CleanUp:
    ' --- 9. 清理物件 ---
    On Error Resume Next ' 忽略可能的清理錯誤
    Set file = Nothing
    Set sourceFolder = Nothing
    Set fso = Nothing
    Set filesToSort = Nothing
    Erase sortedKeys ' 清除陣列
    Exit Sub ' 正常結束

ErrorHandler:
    ' --- 10. 錯誤處理常式 ---
    MsgBox "執行過程中發生未預期的錯誤：" & vbCrLf & vbCrLf & _
           "錯誤代碼: " & Err.Number & vbCrLf & _
           "錯誤描述: " & Err.Description & vbCrLf & _
           "可能發生在檔案: " & IIf(Not file Is Nothing, file.Name, "N/A"), vbCritical, "VBA 執行錯誤"
    Debug.Print "錯誤發生! 代碼: " & Err.Number & ", 描述: " & Err.Description
    Resume CleanUp ' 發生錯誤時，嘗試清理並結束

End Sub

