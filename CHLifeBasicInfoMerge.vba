Option Explicit ' 強制宣告所有變數

Sub Button1_Click()
    Call MergeAndProcessBaseData_V7 ' 更新版本號
End Sub

Sub MergeAndProcessBaseData_V7()

    Dim fso As Object ' FileSystemObject
    Dim sourceFolder As Object
    Dim sourceFile As Object
    Dim folderPath As String
    Dim destWb As Workbook
    Dim sourceWb As Workbook
    Dim sourceSheet As Worksheet
    Dim ws As Worksheet ' 通用工作表變數
    Dim newSheetName As String
    Dim fileCount As Long
    Dim initialSheetCount As Integer ' 記錄新活頁簿初始的工作表數量
    Dim lookupSheet As Worksheet ' 用於查找會所名稱的工作表 (來自唯一的 CSV)
    Dim exportSheet As Worksheet ' 處理中的 export 工作表
    Dim targetSheet As Worksheet ' 最終的 "基本資料" 工作表
    Dim lookupSheetName As String ' 儲存從 CSV 匯入的工作表名稱
    Dim csvFileCount As Integer   ' 計算找到的 CSV 檔案數

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
    lookupSheetName = "" ' 初始化
    csvFileCount = 0     ' 初始化

    ' --- 4. 遍歷資料夾中的檔案 (合併階段) ---
    On Error Resume Next ' 暫時忽略檔案開啟錯誤
    For Each sourceFile In sourceFolder.Files
        Err.Clear
        Dim fileName As String
        Dim fileExt As String
        Dim baseName As String
        fileName = sourceFile.Name
        fileExt = LCase(fso.GetExtensionName(fileName))
        baseName = fso.GetBaseName(sourceFile.Name)

        ' --- 5. 檢查檔案類型 (CSV 或 Export Excel) ---
        Dim isExportExcel As Boolean
        isExportExcel = (InStr(1, fileName, "_export", vbTextCompare) > 0 And _
                         (fileExt = "xlsx" Or fileExt = "xls" Or fileExt = "xlsm"))

        ' --- 6a. 處理 CSV 檔案 (查找表來源) ---
        If fileExt = "csv" Then
            csvFileCount = csvFileCount + 1
            Debug.Print "找到 CSV 檔案: " & fileName

            ' 檢查是否找到多個 CSV
            If csvFileCount > 1 Then
                MsgBox "錯誤：在資料夾中找到超過一個 CSV 檔案。" & vbCrLf & _
                       "請確保資料夾中只有一個 CSV 檔案作為查找表來源。", vbCritical
                GoTo CleanUp ' 或者可以考慮關閉 destWb
            End If

            ' 開啟 CSV
            Set sourceWb = Nothing
            Set sourceWb = Workbooks.Open(sourceFile.Path, ReadOnly:=False, Format:=xlDelimited, Delimiter:=",")

            If Err.Number = 0 And Not sourceWb Is Nothing Then
                If sourceWb.Sheets.Count > 0 Then
                    Set sourceSheet = sourceWb.Sheets(1)
                    newSheetName = CleanSheetName(baseName) ' 清理名稱
                    ' 處理重名
                    Dim counter As Integer
                    Dim tempName As String
                    tempName = newSheetName
                    counter = 1
                    Do While WorksheetExists(destWb, tempName)
                        tempName = Left(newSheetName, 31 - Len(CStr(counter)) - 1) & "_" & CStr(counter)
                        counter = counter + 1
                    Loop
                    newSheetName = tempName

                    sourceSheet.Copy After:=destWb.Sheets(destWb.Sheets.Count)
                    destWb.Sheets(destWb.Sheets.Count).Name = newSheetName
                    lookupSheetName = newSheetName ' *** 儲存此 CSV 匯入後的工作表名稱 ***
                    fileCount = fileCount + 1
                    Debug.Print "  已複製 CSV 工作表，命名為: " & newSheetName & " (將作為查找表)"
                Else
                    Debug.Print "  CSV 檔案 '" & fileName & "' 沒有工作表可複製。"
                End If
                sourceWb.Close SaveChanges:=False
            Else
                Debug.Print "無法開啟或處理 CSV 檔案: " & sourceFile.Path & " - 錯誤: " & Err.Description
                Err.Clear
            End If
            Set sourceSheet = Nothing
            Set sourceWb = Nothing

        ' --- 6b. 處理 Export Excel 檔案 ---
        ElseIf isExportExcel Then
            Debug.Print "找到 Export Excel 檔案: " & fileName
            Set sourceWb = Nothing
            Set sourceWb = Workbooks.Open(sourceFile.Path, ReadOnly:=False)

            If Err.Number = 0 And Not sourceWb Is Nothing Then
                If sourceWb.Sheets.Count > 0 Then
                    Set sourceSheet = sourceWb.Sheets(1)
                    newSheetName = CleanSheetName(baseName)
                    ' 處理重名
                    'Dim counter As Integer
                    'Dim tempName As String
                    tempName = newSheetName
                    counter = 1
                    Do While WorksheetExists(destWb, tempName)
                        tempName = Left(newSheetName, 31 - Len(CStr(counter)) - 1) & "_" & CStr(counter)
                        counter = counter + 1
                    Loop
                    newSheetName = tempName

                    sourceSheet.Copy After:=destWb.Sheets(destWb.Sheets.Count)
                    destWb.Sheets(destWb.Sheets.Count).Name = newSheetName
                    fileCount = fileCount + 1
                    Debug.Print "  已複製 Export 工作表，命名為: " & newSheetName
                Else
                    Debug.Print "  Export Excel 檔案 '" & fileName & "' 沒有工作表可複製。"
                End If
                sourceWb.Close SaveChanges:=False
            Else
                Debug.Print "無法開啟或處理 Export Excel 檔案: " & sourceFile.Path & " - 錯誤: " & Err.Description
                Err.Clear
            End If
            Set sourceSheet = Nothing
            Set sourceWb = Nothing
        Else
            Debug.Print "跳過檔案 (非 CSV 或 Export Excel): " & fileName
        End If
    Next sourceFile
    On Error GoTo 0 ' 恢復正常錯誤處理

    ' --- 7. 清理目標活頁簿的預設工作表 ---
    If destWb.Sheets.Count > initialSheetCount Then
        Dim i As Integer
        Application.DisplayAlerts = False
        For i = initialSheetCount To 1 Step -1
             destWb.Sheets(i).Delete
        Next i
        Application.DisplayAlerts = True
    ElseIf fileCount = 0 Then
        destWb.Close SaveChanges:=False
        MsgBox "在指定的資料夾中沒有找到任何 *.csv 或 *_export*.xls* 檔案。", vbInformation
        GoTo CleanUp
    End If

    ' --- 8. 後處理 ---

    ' --- 8a. 定位查找表 (來自 CSV) 並處理 Export 表格 (添加會所編號) ---
    Dim origFileNameCol As Variant
    Dim huosuoNameCol As Variant
    Dim foundLookupSheet As Boolean
    Dim hallNumber As Variant
    Dim exportSuffix As String
    Dim lastRow As Long
    Dim headerRow As Range

    Set lookupSheet = Nothing
    foundLookupSheet = False

    ' 檢查是否找到了唯一的 CSV 並記錄了其工作表名稱
    Debug.Print "--- 開始定位查找表 (基於 CSV 檔案) ---"
    If csvFileCount = 0 Then
        MsgBox "錯誤：在來源資料夾中未找到任何 CSV 檔案。" & vbCrLf & _
               "需要一個 CSV 檔案作為查找表來源。", vbCritical
        GoTo CleanUp
    ElseIf lookupSheetName = "" Then
         MsgBox "錯誤：找到了 CSV 檔案，但在匯入過程中未能成功記錄其工作表名稱。", vbCritical
         GoTo CleanUp
    End If

    ' 嘗試設定 lookupSheet 物件
    On Error Resume Next
    Set lookupSheet = destWb.Sheets(lookupSheetName)
    On Error GoTo 0

    If lookupSheet Is Nothing Then
        MsgBox "錯誤：無法在活頁簿中找到名為 '" & lookupSheetName & "' 的工作表 (應來自 CSV 檔案)。", vbCritical
        GoTo CleanUp
    End If

    Debug.Print "已定位查找表工作表: '" & lookupSheet.Name & "'"

    ' 檢查查找表標頭
    Set headerRow = lookupSheet.Rows(1)
    On Error Resume Next
    origFileNameCol = Application.Match("原始檔名", headerRow, 0)
    huosuoNameCol = Application.Match("會所名稱", headerRow, 0)
    On Error GoTo 0

    If IsError(origFileNameCol) Or IsError(huosuoNameCol) Then
        MsgBox "錯誤：在查找表工作表 '" & lookupSheet.Name & "' (來自 CSV) 的第一行中，" & vbCrLf & _
               "未能同時找到 '原始檔名' 和 '會所名稱' 這兩個標頭。", vbCritical
        GoTo CleanUp
    Else
        foundLookupSheet = True ' 標頭檢查通過
         Debug.Print "  查找表標頭 ('原始檔名', '會所名稱') 驗證通過。" & _
                      " '原始檔名' 欄位: " & origFileNameCol & _
                      ", '會所名稱' 欄位: " & huosuoNameCol
    End If
    Debug.Print "--- 查找表定位與驗證結束 ---"


    ' 遍歷 Export 工作表並添加會所編號 (此部分邏輯不變)
    If foundLookupSheet Then
        Debug.Print "--- 開始添加會所編號到 Export 工作表 ---"
        For Each exportSheet In destWb.Worksheets
            ' 檢查是否是 Export 表，並且不是查找表本身
            If InStr(1, exportSheet.Name, "_export", vbTextCompare) > 0 And exportSheet.Name <> lookupSheet.Name Then
                Debug.Print "處理 Export 工作表: " & exportSheet.Name
                exportSuffix = GetExportSuffix(exportSheet.Name)
                Debug.Print "  提取的 Export 後綴: '" & exportSuffix & "'"

                If Len(exportSuffix) > 0 Then
                    Dim foundRow As Variant
                    On Error Resume Next
                    foundRow = Application.Match(exportSuffix, lookupSheet.Columns(origFileNameCol), 0)
                    On Error GoTo 0

                    If Not IsError(foundRow) Then
                        Dim hallName As String
                        hallName = CStr(lookupSheet.Cells(foundRow, huosuoNameCol).Value)
                        hallNumber = ConvertToHallNumber(hallName)
                        Debug.Print "    在查找表第 " & foundRow & " 行找到匹配。會所名稱: '" & hallName & "', 轉換結果: " & CStr(hallNumber)

                        exportSheet.Columns(1).Insert Shift:=xlToRight
                        exportSheet.Cells(1, 1).Value = "會所編號"

                        If exportSheet.Cells(exportSheet.Rows.Count, "B").End(xlUp).Row = 1 And IsEmpty(exportSheet.Cells(1, "B").Value) Then
                            lastRow = 1
                        Else
                            lastRow = exportSheet.Cells(exportSheet.Rows.Count, "B").End(xlUp).Row
                        End If

                        If lastRow >= 2 Then
                             If IsError(hallNumber) Then
                                 exportSheet.Range("A2:A" & lastRow).Value = "轉換錯誤"
                                 Debug.Print "      會所名稱轉換失敗，A欄填入 '轉換錯誤'"
                             Else
                                 exportSheet.Range("A2:A" & lastRow).Value = hallNumber
                                 Debug.Print "      已填入會所編號 " & hallNumber & " 到 A2:A" & lastRow
                             End If
                        Else
                            Debug.Print "      工作表 '" & exportSheet.Name & "' 沒有資料列 (只有標題)，不填寫會所編號。"
                        End If
                    Else
                        Debug.Print "    **未找到** 匹配的 Export 後綴: '" & exportSuffix & "'"
                        exportSheet.Columns(1).Insert Shift:=xlToRight
                        exportSheet.Cells(1, 1).Value = "會所編號(未找到)"
                    End If
                Else
                     Debug.Print "  **警告**: 無法從工作表名稱 '" & exportSheet.Name & "' 提取 Export 後綴 (_export...)"
                     exportSheet.Columns(1).Insert Shift:=xlToRight
                     exportSheet.Cells(1, 1).Value = "會所編號(名稱錯誤)"
                End If
            End If
        Next exportSheet
        Debug.Print "--- 會所編號添加完成 ---"
    Else
         Debug.Print "由於查找表未通過驗證，跳過添加會所編號步驟。"
         ' 這裡已經 GoTo CleanUp 了，理論上不會執行到
    End If


    ' --- 8b. 創建 "基本資料" 工作表並合併數據 (邏輯不變) ---
    Debug.Print "--- 開始創建 '基本資料' 並合併數據 ---"
    Set targetSheet = Nothing
    On Error Resume Next
    Set targetSheet = destWb.Sheets("基本資料")
    On Error GoTo 0
    If Not targetSheet Is Nothing Then
        Application.DisplayAlerts = False
        targetSheet.Delete
        Application.DisplayAlerts = True
        Debug.Print "已刪除已存在的 '基本資料' 工作表。"
    End If

    Set targetSheet = destWb.Sheets.Add(After:=destWb.Sheets(destWb.Sheets.Count))
    targetSheet.Name = "基本資料"
    Debug.Print "已創建新的工作表 '基本資料'"

    Dim headerCopied As Boolean
    Dim destNextRow As Long
    Dim sourceLastRow As Long
    Dim sourceLastCol As Long
    Dim firstExportSheet As Worksheet

    headerCopied = False
    destNextRow = 1

    Set firstExportSheet = Nothing
    For Each ws In destWb.Worksheets
        If InStr(1, ws.Name, "_export", vbTextCompare) > 0 And ws.Name <> lookupSheet.Name And ws.Name <> targetSheet.Name Then
             If Not IsEmpty(ws.Cells(1, 1).Value) Then
                 Set firstExportSheet = ws
                 Exit For
             End If
        End If
    Next ws

    If Not firstExportSheet Is Nothing Then
        On Error Resume Next
        firstExportSheet.Rows(1).Copy Destination:=targetSheet.Rows(1)
        If Err.Number = 0 Then
            headerCopied = True
            destNextRow = 2
            Debug.Print "已從 '" & firstExportSheet.Name & "' 複製標題到 '基本資料'"
        Else
            Debug.Print "**錯誤**: 無法從 '" & firstExportSheet.Name & "' 複製標題. Error: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Else
        Debug.Print "**警告**: 未找到任何已處理的 Export 工作表來複製標題。"
    End If

    If headerCopied Then
        For Each ws In destWb.Worksheets
            If InStr(1, ws.Name, "_export", vbTextCompare) > 0 And ws.Name <> lookupSheet.Name And ws.Name <> targetSheet.Name Then
                Debug.Print "  合併來自 '" & ws.Name & "' 的數據..."
                sourceLastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
                If sourceLastRow >= 2 Then
                     sourceLastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
                     If sourceLastCol < 1 Then sourceLastCol = 1

                     Dim sourceDataRange As Range
                     Set sourceDataRange = ws.Range(ws.Cells(2, 1), ws.Cells(sourceLastRow, sourceLastCol))

                     On Error Resume Next
                     sourceDataRange.Copy Destination:=targetSheet.Cells(destNextRow, 1)
                     If Err.Number = 0 Then
                         Debug.Print "    已複製範圍 " & ws.Name & "!" & sourceDataRange.Address & " 到 '基本資料' 第 " & destNextRow & " 行"
                         destNextRow = targetSheet.Cells(targetSheet.Rows.Count, "A").End(xlUp).Row + 1
                     Else
                         Debug.Print "    **錯誤**: 複製範圍 " & ws.Name & "!" & sourceDataRange.Address & " 失敗. Error: " & Err.Description
                         Err.Clear
                     End If
                     Set sourceDataRange = Nothing
                     On Error GoTo 0
                     Application.CutCopyMode = False
                Else
                    Debug.Print "    工作表 '" & ws.Name & "' 沒有數據行 (第2行及以後) 可合併。"
                End If
            End If
        Next ws
    End If
    Debug.Print "--- 數據合併完成 ---"

    ' --- 8c. 刪除原始工作表，只保留 "基本資料" (邏輯不變) ---
    Debug.Print "--- 開始刪除原始工作表 ---"
    Application.DisplayAlerts = False
    Dim sheetToDelete As Worksheet
    For i = destWb.Worksheets.Count To 1 Step -1
        Set sheetToDelete = destWb.Worksheets(i)
        If sheetToDelete.Name <> targetSheet.Name Then
            Debug.Print "  刪除工作表: " & sheetToDelete.Name
            On Error Resume Next
            sheetToDelete.Delete
            If Err.Number <> 0 Then
                Debug.Print "    **錯誤**: 刪除工作表 '" & sheetToDelete.Name & "' 失敗. Error: " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next i
    Application.DisplayAlerts = True
    Debug.Print "--- 原始工作表刪除完成 ---"

    On Error Resume Next
    targetSheet.Columns.AutoFit
    On Error GoTo 0


CleanUp:
    ' --- 9. 收尾 ---
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    Set sourceFile = Nothing
    Set sourceFolder = Nothing
    Set fso = Nothing
    If Not destWb Is Nothing Then
        ' 如果 lookupSheet 未成功找到或驗證失敗，可能需要決定是否關閉 destWb
        If Not foundLookupSheet And fileCount > 0 And csvFileCount = 1 Then
             ' 可能需要提示使用者查找表有問題，並保留活頁簿供檢查
             MsgBox "處理中止。請檢查活頁簿中從 CSV 檔案匯入的工作表 '" & lookupSheetName & "' 的標頭是否正確。", vbExclamation
        ElseIf Not Application.Calculation = xlCalculationManual Then ' 避免在完全失敗時關閉
            ' 正常情況下保持活頁簿開啟
        Else ' 如果因多個 CSV 或完全沒檔案而跳到 CleanUp，可能需要關閉
            If csvFileCount > 1 Or fileCount = 0 Then
                 If Not destWb Is Nothing Then destWb.Close SaveChanges:=False
            End If
        End If
    End If
    ' Set destWb = Nothing ' 取決於上面邏輯

    Set sourceWb = Nothing
    Set sourceSheet = Nothing
    Set ws = Nothing
    Set lookupSheet = Nothing
    Set exportSheet = Nothing
    Set targetSheet = Nothing
    Set headerRow = Nothing
    Set firstExportSheet = Nothing
    Set sheetToDelete = Nothing


    ' --- 10. 完成提示 ---
    If Not destWb Is Nothing And Application.Calculation = xlCalculationAutomatic Then ' 確保是在正常完成時提示
        If fileCount > 0 And foundLookupSheet Then
            MsgBox "合併與處理完成！共處理了 " & fileCount & " 個檔案 (含 1 個 CSV 查找表)。" & vbCrLf & _
                   "所有數據已合併到 '基本資料' 工作表中。", vbInformation
        ' 其他錯誤情況的 MsgBox 已在前面處理
        End If
    End If


End Sub

'========================================================================================
' Helper Function: CleanSheetName (清理工作表名稱，保持不變)
' ... (代碼省略) ...
'========================================================================================
Private Function CleanSheetName(ByVal sheetName As String) As String
    Dim invalidChars As String
    Dim i As Integer
    Dim tempChar As String

    sheetName = Trim(sheetName)
    If Len(sheetName) = 0 Then
        CleanSheetName = "_"
        Exit Function
    End If

    invalidChars = "\/?*[]:" ' Excel 工作表名稱不允許的字元

    For i = 1 To Len(invalidChars)
        tempChar = Mid(invalidChars, i, 1)
        sheetName = Replace(sheetName, tempChar, "_")
    Next i

    ' 工作表名稱長度限制為 31
    If Len(sheetName) > 31 Then
        sheetName = Left(sheetName, 31)
    End If

    ' 工作表名稱不能以單引號開頭或結尾
    If Left(sheetName, 1) = "'" Then sheetName = "_" & Mid(sheetName, 2)
    If Right(sheetName, 1) = "'" Then sheetName = Mid(sheetName, 1, Len(sheetName) - 1) & "_"

    ' 如果清理後變成空字串，給個預設值
    If Len(Trim(sheetName)) = 0 Then sheetName = "_"

    CleanSheetName = sheetName
End Function

'========================================================================================
' Helper Function: GetExportSuffix (提取 Export 後綴，保持不變)
' ... (代碼省略) ...
'========================================================================================
Private Function GetExportSuffix(sheetName As String) As String
    Dim position As Long
    Const SUFFIX_START As String = "_export"

    ' 查找 "_export" 的位置 (不區分大小寫)
    position = InStr(1, sheetName, SUFFIX_START, vbTextCompare)

    If position > 0 Then
        ' 提取從 "_export" 開始到結尾的部分
        GetExportSuffix = Mid(sheetName, position)
    Else
        ' 如果找不到 "_export"，返回空字串
        GetExportSuffix = ""
        ' Debug.Print "**警告**: 在 '" & sheetName & "' 中未找到 '" & SUFFIX_START & "'，無法提取 Export 後綴。"
    End If
End Function

'========================================================================================
' Helper Function: WorksheetExists (檢查工作表是否存在，保持不變)
' ... (代碼省略) ...
'========================================================================================
Private Function WorksheetExists(wb As Workbook, sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    WorksheetExists = Not ws Is Nothing
End Function


'========================================================================================
' Main Function: ConvertToHallNumber (從會所名稱提取數字，保持不變)
' ... (代碼省略) ...
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
        chineseNumStr = Trim(Mid(inputText, startPos + Len(START_DELIMITER), endPos - startPos - Len(START_DELIMITER)))

        ' Try to convert the extracted Chinese numeral string to an Arabic number
        On Error Resume Next ' Enable error trapping for the conversion function
        result = ChineseToArabic(chineseNumStr)
        On Error GoTo 0 ' Disable error trapping

        ' Check if the conversion function returned an error
        If IsError(result) Then
            'ConvertToHallNumber = CVErr(xlErrValue) ' Return #VALUE! if conversion failed
             ConvertToHallNumber = CVErr(xlErrNA) ' 改為回傳 #N/A 可能更清晰表示轉換問題
             Debug.Print "ConvertToHallNumber: Conversion failed for '" & chineseNumStr & "' from '" & inputText & "'"
        Else
            ConvertToHallNumber = result ' Return the converted number
        End If
    Else
        ' Input string does not match the expected format "第...會所"
        ConvertToHallNumber = CVErr(xlErrNA) ' Return #N/A for format mismatch
        Debug.Print "ConvertToHallNumber: Format mismatch for '" & inputText & "'"
    End If

End Function

'========================================================================================
' Helper Function: ChineseToArabic (中文數字轉阿拉伯數字，保持不變)
' ... (代碼省略) ...
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
    ' Handle special case: 十 at the beginning (e.g., 十二) -> 1十
    If Left(tempStr, 1) = "十" Then tempStr = "1" & tempStr


    ' --- Step 2: Parse the mixed string (digits, 百, 十) ---
    On Error GoTo ConversionError ' Trap errors during Val() or logic

    resultVal = 0 ' Initialize result

    ' --- Handle Hundreds (百) ---
    posBai = InStr(1, tempStr, "百", vbTextCompare)
    If posBai > 0 Then
        Dim leftBai As String
        Dim valBai As Long

        leftBai = Trim(Mid(tempStr, 1, posBai - 1))

        If Len(leftBai) = 0 Then ' Case: "百" (implies 100) - Should not happen after replacements
             GoTo ConversionError ' Invalid format like "百二"
        ElseIf IsNumeric(leftBai) Then
             valBai = Val(leftBai)
        Else
             GoTo ConversionError ' Invalid character before 百
        End If

        resultVal = valBai * 100
        ' Process the remainder of the string after 百 for tens/units
        tempStr = Trim(Mid(tempStr, posBai + 1))
         ' Handle case like "一百" (no tens or units)
        If Len(tempStr) = 0 Then GoTo FinalizeResult ' Nothing left to parse

    End If ' End of handling 百

    ' --- Handle Tens (十) and Units ---
    posShi = InStr(1, tempStr, "十", vbTextCompare)
    If posShi > 0 Then
         Dim leftShi As String, rightShi As String
         Dim valShiTens As Long, valShiUnits As Long

         leftShi = Trim(Mid(tempStr, 1, posShi - 1))
         rightShi = Trim(Mid(tempStr, posShi + 1))

         ' Calculate Tens part
         If Len(leftShi) = 0 Then ' Should not happen due to initial "十" -> "1十" replacement
             GoTo ConversionError ' Invalid format like "十" alone after 百
         ElseIf IsNumeric(leftShi) Then ' Case: "1十" (10), "5十" (50)
             valShiTens = Val(leftShi) * 10
         Else
             GoTo ConversionError ' Invalid character before 十
         End If

         ' Calculate Units part
         If Len(rightShi) = 0 Then ' Case: "五十" (no units)
             valShiUnits = 0
         ElseIf rightShi = "0" Then ' Handle "...十零" case, units are 0
             valShiUnits = 0
         ElseIf IsNumeric(rightShi) Then ' Case: "十二", "五十五" -> "1十2", "5十5"
              valShiUnits = Val(rightShi)
         Else
             GoTo ConversionError ' Invalid character after 十
         End If

         resultVal = resultVal + valShiTens + valShiUnits ' Add tens and units to hundreds (if any)

    Else ' No "十" found in the remaining string (or original string without 百)
        ' Remaining string should be just digits (units or zero after 百, e.g., "一百零五" -> "05")
        If Len(tempStr) > 0 Then
            If IsNumeric(tempStr) Then
                 resultVal = resultVal + Val(tempStr)
            ElseIf tempStr = "0" Then ' Handle case like "一百零" -> result is 100
                 ' Do nothing, resultVal is correct
            Else
                 GoTo ConversionError ' Invalid characters remain
            End If
        End If ' If tempStr is empty here (e.g. after "一百"), resultVal is already correct
    End If ' End of handling 十

FinalizeResult:
    ' Return the final calculated value
    ChineseToArabic = resultVal
    Exit Function

ConversionError:
    ' If any error occurred during parsing (e.g., non-numeric parts)
    Debug.Print "ChineseToArabic Error: Could not parse '" & chinNum & "' (Processed as '" & tempStr & "')"
    ChineseToArabic = CVErr(xlErrValue) ' Return #VALUE!

End Function


'========================================================================================
' Example Subroutine to demonstrate usage (保持不變)
' ... (代碼省略) ...
'========================================================================================
Sub ProcessHallData()
    Dim hallName As String
    Dim hallNumber As Variant

    hallName = "台北市召會第六十會所"
    hallNumber = ConvertToHallNumber(hallName)

    If IsError(hallNumber) Then
        MsgBox "無法轉換會所名稱: " & hallName & vbCrLf & "錯誤代碼: " & CStr(hallNumber)
    Else
        MsgBox "轉換結果: " & hallName & " -> " & CStr(hallNumber)
        ' Range("B1").Value = hallNumber ' Put the result in cell B1
    End If

    hallName = "台北市召會第102會所" ' Test case with Arabic numerals mixed
    hallNumber = ConvertToHallNumber(hallName)
     If IsError(hallNumber) Then
        MsgBox "無法轉換會所名稱: " & hallName & vbCrLf & "錯誤代碼: " & CStr(hallNumber)
    Else
        MsgBox "轉換結果: " & hallName & " -> " & CStr(hallNumber)
    End If

     hallName = "台北市召會第一百零七會所" ' Test case with 百 and 零
    hallNumber = ConvertToHallNumber(hallName)
     If IsError(hallNumber) Then
        MsgBox "無法轉換會所名稱: " & hallName & vbCrLf & "錯誤代碼: " & CStr(hallNumber)
    Else
        MsgBox "轉換結果: " & hallName & " -> " & CStr(hallNumber)
    End If

     hallName = "Some Other Text" ' Test case with wrong format
    hallNumber = ConvertToHallNumber(hallName)
     If IsError(hallNumber) Then
        MsgBox "無法轉換會所名稱: " & hallName & vbCrLf & "錯誤代碼: " & CStr(hallNumber)
    Else
        MsgBox "轉換結果: " & hallName & " -> " & CStr(hallNumber)
    End If

    hallName = "台北市召會第七十會所"
    hallNumber = ConvertToHallNumber(hallName)
     If IsError(hallNumber) Then
        MsgBox "無法轉換會所名稱: " & hallName & vbCrLf & "錯誤代碼: " & CStr(hallNumber)
    Else
        MsgBox "轉換結果: " & hallName & " -> " & CStr(hallNumber)
    End If

End Sub

