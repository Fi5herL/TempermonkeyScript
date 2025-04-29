Sub Button3_Click()
    Call MergeVerificationCodes_V3_HandleYearOnlyDate
End Sub

Sub MergeVerificationCodes_V3_HandleYearOnlyDate()

    Dim wbSource As Workbook ' 基本資料 Workbook
    Dim wbTarget As Workbook ' 聚會資料 Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim dictLookup As Object ' Dictionary for lookup
    Dim sourceFilePath As Variant
    Dim targetFilePath As Variant
    Dim sourceHeaderRow As Long
    Dim targetHeaderRow As Long
    Dim sourceDataStartRow As Long
    Dim targetDataStartRow As Long
    Dim sourceLastRow As Long
    Dim targetLastRow As Long
    Dim r As Long ' Row index
    Dim keyString As String
    Dim keyColsSource As Object ' 儲存來源 Key 欄位的欄號 (Dictionary: Name -> ColNum)
    Dim keyColsTarget As Object ' 儲存目標 Key 欄位的欄號 (Dictionary: Name -> ColNum)
    Dim keyColNamesSource As Variant ' 來源 Key 欄位的名稱 Array
    Dim keyColNamesTarget As Variant ' 目標 Key 欄位的名稱 Array (用於查找欄號)
    Dim keyColMapping As Object ' Dictionary for mapping Target Name -> Source Name for key building
    Dim verificationColName As String
    Dim verificationColNumSource As Long
    Dim huosuoColNameSource As String ' 基本資料的會所編號欄位名
    Dim huosuoColNumSource As Long
    Dim targetHuosuoCol As Long ' 聚會資料的會所編號欄位號 (通常是 1 or "A")
    Dim i As Integer
    Dim colName As Variant
    Dim cellValue As Variant
    Dim keyValues() As String ' Array to hold values for key string
    Dim missingCol As Boolean

    ' --- 設定 ---
    sourceHeaderRow = 1 ' 基本資料的標頭列
    sourceDataStartRow = 2 ' 基本資料的資料起始列
    targetHeaderRow = 2 ' 聚會資料的標頭列 (根據範例)
    targetDataStartRow = 3 ' 聚會資料的資料起始列

    huosuoColNameSource = "會所編號" ' 基本資料中會所編號的欄位名稱
    targetHuosuoCol = 1 ' 聚會資料中會所編號所在的欄位 (通常是第一欄 A) - ***注意：插入前是A(1)，插入後會變B(2)***

    ' *** V6.1 修改：定義新的基本資料日期欄位名稱 ***
    sourceDateFieldColName = "受浸日期 YYYY-MM-DD（可只填年YYYY或年月YYYY-MM，或用『/』分隔）"

    ' 基本資料中用於比對的欄位名稱 (這些是用來從基本資料讀取的)
    keyColNamesSource = Array(huosuoColNameSource, "大區", "小區", "姓名", sourceDateFieldColName)
    ' 聚會資料中對應的欄位名稱 (這些是用來從聚會資料讀取的，順序需與 Source 一致以建立 Key)
    ' 注意：第一個會所編號是直接用欄位號 targetHuosuoCol 讀取，這裡放其他欄位名
    keyColNamesTarget = Array("大區", "小區", "姓名", "受浸日期") ' 會所編號欄位在 Target 端單獨處理

    verificationColName = "驗證碼" ' 基本資料中驗證碼欄位的名稱
    Const DELIMITER As String = "|~|" ' 用於組合 Key 的分隔符號

    ' 建立 Target -> Source 的欄位名稱映射 (用於建立一致的 Key)
    Set keyColMapping = CreateObject("Scripting.Dictionary")
    keyColMapping.CompareMode = vbTextCompare
    keyColMapping.Add "大區", "大區"
    keyColMapping.Add "小區", "小區"
    keyColMapping.Add "姓名", "姓名"
    keyColMapping.Add "受浸日期", "受浸日期"

    ' --- 建立字典物件 ---
    Set dictLookup = CreateObject("Scripting.Dictionary")
    dictLookup.CompareMode = vbTextCompare ' 設定為不區分大小寫比對

    ' --- 關閉螢幕更新以加速 ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
    Application.StatusBar = "正在初始化..."

    ' --- 1. 選擇並開啟/設定基本資料檔案 ---
    MsgBox "請選擇 '基本資料' Excel 檔案", vbInformation
    sourceFilePath = Application.GetOpenFilename("Excel Files (*.xlsx; *.xls; *.xlsm),*.xlsx;*.xls;*.xlsm", , "請選擇基本資料檔案")
    If sourceFilePath = False Then GoTo CleanUpEarly: MsgBox "未選擇基本資料檔案，程序中止。", vbExclamation: GoTo CleanUp
    On Error Resume Next
    Set wbSource = Workbooks(Dir(sourceFilePath))
    On Error GoTo 0
    If wbSource Is Nothing Then Set wbSource = Workbooks.Open(sourceFilePath)
    If wbSource Is Nothing Then GoTo CleanUpEarly: MsgBox "無法開啟基本資料檔案 '" & sourceFilePath & "'。", vbCritical: GoTo CleanUp
    Debug.Print "基本資料檔案: " & wbSource.Name
    Application.StatusBar = "正在讀取基本資料並建立查找字典..."

    ' --- 2. 讀取基本資料並建立查找字典 ---
    Debug.Print "正在建立查找字典..."
    Set keyColsSource = CreateObject("Scripting.Dictionary") ' 使用字典存儲 Name -> ColNum

    For Each wsSource In wbSource.Worksheets
        Debug.Print "  處理基本資料分頁: " & wsSource.Name
        Application.StatusBar = "處理基本資料分頁: " & wsSource.Name & "..."
        keyColsSource.RemoveAll ' 清空上一頁的欄位資訊
        missingCol = False

        ' 尋找所有需要的來源欄位號碼
        For Each colName In keyColNamesSource
            Dim colNum As Long
            colNum = GetColumnNumber(wsSource, sourceHeaderRow, CStr(colName))
            If colNum = 0 Then
                Debug.Print "    警告: 在分頁 '" & wsSource.Name & "' 找不到欄位 '" & colName & "'，跳過此分頁。"
                ' MsgBox "警告: 在基本資料分頁 '" & wsSource.Name & "' 找不到欄位 '" & colName & "'，將跳過此分頁。", vbExclamation ' 暫時關閉彈窗提示
                missingCol = True
                Exit For ' 缺少一個就跳過此分頁
            End If
            keyColsSource.Add CStr(colName), colNum ' 儲存 Name -> ColNum
        Next colName
        If missingCol Then GoTo NextSourceSheet ' 跳到下一個分頁

        ' 尋找驗證碼欄位號碼
        verificationColNumSource = GetColumnNumber(wsSource, sourceHeaderRow, verificationColName)
        If verificationColNumSource = 0 Then
            Debug.Print "    警告: 在分頁 '" & wsSource.Name & "' 找不到欄位 '" & verificationColName & "'，跳過此分頁。"
            ' MsgBox "警告: 在基本資料分頁 '" & wsSource.Name & "' 找不到欄位 '" & verificationColName & "'，將跳過此分頁。", vbExclamation ' 暫時關閉彈窗提示
            GoTo NextSourceSheet
        End If

        ' 尋找最後一列 (以會所編號欄位為基準)
        huosuoColNumSource = keyColsSource(huosuoColNameSource)
        sourceLastRow = wsSource.Cells(wsSource.Rows.Count, huosuoColNumSource).End(xlUp).Row

        ' 遍歷資料列
        ReDim keyValues(LBound(keyColNamesSource) To UBound(keyColNamesSource)) ' Dimension the array
        For r = sourceDataStartRow To sourceLastRow
            On Error Resume Next ' 忽略單一儲存格讀取錯誤
            For i = LBound(keyColNamesSource) To UBound(keyColNamesSource)
                colName = keyColNamesSource(i) ' Get the current source column name
                cellValue = wsSource.Cells(r, keyColsSource(CStr(colName))).Value

                ' *** V5 修改：調用新的日期格式化函數 ***
                If CStr(colName) = sourceDateFieldColName Then
                    keyValues(i) = FormatBaptismDateForKey(cellValue) ' 使用 V5 函數
                Else
                    keyValues(i) = CleanValue(cellValue)
                End If
                ' *** V5 修改結束 ***
            Next i
            On Error GoTo 0 ' 恢復錯誤處理

            keyString = Join(keyValues, DELIMITER) ' 組合 Key

            ' 獲取驗證碼並清理
            Dim verificationCode As String
            verificationCode = CleanValue(wsSource.Cells(r, verificationColNumSource).Value)

            ' 添加到字典
            If Not dictLookup.Exists(keyString) Then
                dictLookup.Add keyString, verificationCode
            Else
                dictLookup(keyString) = verificationCode ' 更新重複的 Key
            End If

            If r Mod 500 = 0 Then ' 每 500 行更新一次狀態
                Application.StatusBar = "處理基本資料分頁: " & wsSource.Name & " (Row " & r & "/" & sourceLastRow & ")..."
                DoEvents
            End If
        Next r
NextSourceSheet:
    Next wsSource
    Debug.Print "查找字典建立完成，共 " & dictLookup.Count & " 筆記錄。"
    Application.StatusBar = "查找字典建立完成。請選擇聚會資料檔案..."

    ' --- 3. 選擇並開啟/設定聚會資料檔案 ---
    MsgBox "請選擇 '聚會資料' Excel 檔案", vbInformation
    targetFilePath = Application.GetOpenFilename("Excel Files (*.xlsx; *.xls; *.xlsm),*.xlsx;*.xls;*.xlsm", , "請選擇聚會資料檔案")
    If targetFilePath = False Then GoTo CleanUpEarly: MsgBox "未選擇聚會資料檔案，程序中止。", vbExclamation: GoTo CleanUp
    On Error Resume Next
    Set wbTarget = Workbooks(Dir(targetFilePath))
    On Error GoTo 0
    If wbTarget Is Nothing Then Set wbTarget = Workbooks.Open(targetFilePath)
    If wbTarget Is Nothing Then GoTo CleanUpEarly: MsgBox "無法開啟聚會資料檔案 '" & targetFilePath & "'。", vbCritical: GoTo CleanUp
    Debug.Print "聚會資料檔案: " & wbTarget.Name
    Application.StatusBar = "正在處理聚會資料..."

    ' --- 4. 處理聚會資料 ---
    Debug.Print "正在處理聚會資料..."
    Set keyColsTarget = CreateObject("Scripting.Dictionary") ' Name -> ColNum

    For Each wsTarget In wbTarget.Worksheets
        Debug.Print "  處理聚會資料分頁: " & wsTarget.Name
        Application.StatusBar = "處理聚會資料分頁: " & wsTarget.Name & "..."
        keyColsTarget.RemoveAll
        missingCol = False

        ' --- 預處理：插入欄位 ---
        ' 檢查 A 欄是否已存在 "驗證碼"，如果存在則先刪除 (避免重複執行時不斷插入)
        If LCase(Trim(CStr(wsTarget.Cells(targetHeaderRow, "A").Value))) = LCase(verificationColName) Then
             wsTarget.Columns("A").Delete Shift:=xlToLeft
             Debug.Print "    發現已存在的 '驗證碼' 欄位，已刪除。"
        End If
        ' 插入新欄位在 A 欄
        wsTarget.Columns("A").Insert Shift:=xlToRight
        wsTarget.Cells(targetHeaderRow, "A").Value = verificationColName ' 寫入標頭
        Debug.Print "    已插入 '驗證碼' 欄位。"
        ' --- 預處理結束 ---

        ' 尋找目標欄位號碼 (在插入新欄後的位置)
        ' 會所編號現在是 B 欄 (欄號 2)
        targetHuosuoCol = 2 ' *** 插入新欄後，原來的 A 欄變為 B 欄 ***

        ' 尋找其他 Key 欄位
        For Each colName In keyColNamesTarget
            ' Dim colNum As Long
            colNum = GetColumnNumber(wsTarget, targetHeaderRow, CStr(colName))
            If colNum = 0 Then
                Debug.Print "    警告: 在分頁 '" & wsTarget.Name & "' 找不到欄位 '" & colName & "' (插入新欄後)，跳過此分頁。"
                ' MsgBox "警告: 在聚會資料分頁 '" & wsTarget.Name & "' 找不到欄位 '" & colName & "'，將跳過此分頁的資料處理。", vbExclamation ' 暫時關閉彈窗提示
                missingCol = True
                Exit For
            End If
            keyColsTarget.Add CStr(colName), colNum ' 儲存 Name -> ColNum
        Next colName
        If missingCol Then GoTo NextTargetSheet ' 跳到下一個目標分頁

        ' 尋找最後一列 (以插入欄後的 B 欄，即原來的會所編號欄為基準)
        targetLastRow = wsTarget.Cells(wsTarget.Rows.Count, targetHuosuoCol).End(xlUp).Row
        Debug.Print "    處理列數: " & targetLastRow - targetDataStartRow + 1

        ' 遍歷資料列 (從資料起始列開始)
        Dim foundCount As Long
        Dim notFoundCount As Long
        foundCount = 0
        notFoundCount = 0
        ReDim keyValues(LBound(keyColNamesSource) To UBound(keyColNamesSource)) ' 重用 keyValues 陣列，大小與 Source Key 一致

        For r = targetDataStartRow To targetLastRow
            On Error Resume Next ' 忽略單一儲存格讀取錯誤
            ' 按照 Source Key 的順序建立 Target Key
            ' 1. 會所編號 (從 B 欄讀取)
            keyValues(0) = CleanValue(wsTarget.Cells(r, targetHuosuoCol).Value)

            ' 2. 其他欄位 (使用 keyColNamesTarget 找到的欄號，並按 Source 順序填充)
            Dim targetIdx As Long
            targetIdx = 1 ' keyValues 陣列的索引從 1 開始對應其他欄位
            For Each colName In keyColNamesTarget ' 遍歷 Target 欄位名
                ' Dim sourceMappedName As String ' 不需要讀取 Source Name 了
                ' sourceMappedName = keyColMapping(CStr(colName))

                ' 找到此 Target 欄位在 wsTarget 中的欄號
                Dim currentTargetColNum As Long
                currentTargetColNum = keyColsTarget(CStr(colName))

                ' 讀取並清理值
                cellValue = wsTarget.Cells(r, currentTargetColNum).Value

                ' *** V5 修改：調用新的日期格式化函數 ***
                If CStr(colName) = "受浸日期" Then
                    keyValues(targetIdx) = FormatBaptismDateForKey(cellValue) ' 使用新函數處理日期
                Else ' 其他欄位使用通用清理函數
                    keyValues(targetIdx) = CleanValue(cellValue)
                End If
                ' *** V5 修改結束 ***

                targetIdx = targetIdx + 1
            Next colName
            On Error GoTo 0 ' 恢復錯誤處理

            keyString = Join(keyValues, DELIMITER) ' 組合 Target Key

            ' 在字典中查找 Key
            If dictLookup.Exists(keyString) Then
                wsTarget.Cells(r, "A").Value = dictLookup(keyString) ' 填入驗證碼到 A 欄
                foundCount = foundCount + 1
            Else
                wsTarget.Cells(r, "A").Value = "" ' 找不到則留空
                notFoundCount = notFoundCount + 1
                 ' If r < targetDataStartRow + 20 And notFoundCount < 20 Then ' 只打印前幾筆找不到的 Key 以供除錯
                 '    Debug.Print "    Key not found: """ & keyString & """ (Row " & r & ", Sheet: " & wsTarget.Name & ")"
                 ' End If
            End If

            If r Mod 500 = 0 Then ' 每 500 行更新一次狀態
                Application.StatusBar = "處理聚會資料分頁: " & wsTarget.Name & " (Row " & r & "/" & targetLastRow & ")..."
                DoEvents
            End If
        Next r
        Debug.Print "    分頁 '" & wsTarget.Name & "' 完成: 找到 " & foundCount & " 筆, 未找到 " & notFoundCount & " 筆。"
NextTargetSheet:
        On Error Resume Next ' 忽略 AutoFit 可能的錯誤 (例如隱藏欄)
        wsTarget.Columns("A:" & wsTarget.UsedRange.Address).AutoFit ' 只自動調整使用的欄寬
        On Error GoTo 0
    Next wsTarget

    ' --- 儲存修改後的聚會資料檔案 ---
    On Error Resume Next
    wbTarget.Save
    If Err.Number <> 0 Then
        MsgBox "儲存聚會資料檔案 '" & wbTarget.Name & "' 時發生錯誤。" & vbCrLf & _
               "錯誤描述: " & Err.Description & vbCrLf & vbCrLf & _
               "請手動儲存該檔案。", vbExclamation
        Err.Clear
    Else
        Debug.Print "聚會資料檔案已儲存。"
    End If
    On Error GoTo 0

    ' --- 完成提示 ---
    Application.StatusBar = "處理完成！"
    MsgBox "處理完成！驗證碼已新增至 '" & wbTarget.Name & "' 的每個分頁的第一欄。", vbInformation

CleanUp:
    ' --- 恢復設定 & 清理物件 ---
    On Error Resume Next ' 忽略可能的關閉錯誤
    Set dictLookup = Nothing
    Set keyColsSource = Nothing
    Set keyColsTarget = Nothing
    Set keyColMapping = Nothing
    Set wsSource = Nothing
    Set wsTarget = Nothing
    ' If Not wbSource Is Nothing And UCase(wbSource.FullName) = UCase(sourceFilePath) Then wbSource.Close SaveChanges:=False
    Set wbSource = Nothing
    Set wbTarget = Nothing
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Debug.Print "清理完成。"
    Exit Sub ' 正常退出

CleanUpEarly: ' 如果選擇檔案階段就出錯
    Application.StatusBar = False
    GoTo CleanUp

End Sub

' --- Helper Function: 尋找欄位名稱對應的欄號 ---
Function GetColumnNumber(ws As Worksheet, headerRow As Long, colName As String) As Long
    Dim foundCell As Range
    Dim searchRange As Range
    GetColumnNumber = 0 ' Default to not found

    If headerRow <= 0 Then Exit Function ' Invalid header row
    ' 限制搜尋範圍，避免對整行操作，提高效率
    On Error Resume Next
    Set searchRange = ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, ws.UsedRange.Columns.Count))
    If Err.Number <> 0 Then Set searchRange = ws.Rows(headerRow) ' Fallback if UsedRange fails
    On Error GoTo 0
    If searchRange Is Nothing Then Exit Function

    On Error Resume Next
    Set foundCell = searchRange.Find(What:=Trim(colName), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False, SearchFormat:=False)
    On Error GoTo 0

    If Not foundCell Is Nothing Then
        GetColumnNumber = foundCell.Column
    End If
    Set foundCell = Nothing
    Set searchRange = Nothing
End Function

' --- Helper Function: 清理儲存格值以進行比較 ---
Function CleanValue(cellValue As Variant) As String
    Dim cleaned As String
    On Error Resume Next ' Handle potential errors during conversion/formatting

    If IsError(cellValue) Then
        cleaned = "#ERROR!"
    ElseIf IsDate(cellValue) Then
        Dim dt As Date
        dt = CDate(cellValue)
        If Year(dt) > 1900 And Month(dt) > 0 And Day(dt) > 0 Then
             cleaned = Format(dt, "yyyy-mm-dd")
        ElseIf Year(dt) > 1900 And Month(dt) > 0 Then
             cleaned = Format(dt, "yyyy-mm")
        ElseIf Year(dt) > 1900 Then
             cleaned = Format(dt, "yyyy") ' Keep only year if month/day are 0
        Else
             cleaned = Trim(CStr(cellValue)) ' Fallback for unusual dates
        End If
        ' Handle specific cases like "0000-..." which might parse strangely
        If InStr(cleaned, "1899") > 0 Or (InStr(cleaned, "1900") > 0 And LCase(Trim(CStr(cellValue))) Like "*00*") Then
             cleaned = Trim(CStr(cellValue))
        End If
    ElseIf IsNumeric(cellValue) Then
         cleaned = Trim(CStr(cellValue)) ' Convert numbers to string
    Else
        cleaned = Trim(CStr(cellValue)) ' Convert other types to string
    End If
    On Error GoTo 0

    cleaned = Replace(cleaned, ChrW(160), "") ' Replace non-breaking space
    CleanValue = Trim(cleaned) ' Final trim

End Function

' --- Helper Function: 格式化受浸日期用於 Key 比對 (版本 V6 - 統一補零) ---
Function FormatBaptismDateForKey(cellValue As Variant) As String
    Dim rawValue As String
    Dim yearPart As String, monthPart As String, dayPart As String
    Dim parts() As String

    ' 1. 初始處理：處理 Null 和空字串
    If IsNull(cellValue) Or Trim(CStr(cellValue)) = "" Then
        FormatBaptismDateForKey = "0000-00-00"
        Exit Function
    End If

    ' 2. 轉換為字串並清理
    rawValue = Trim(CStr(cellValue))

    ' 3. 直接處理 "0000-00-00"
    If rawValue = "0000-00-00" Then
        FormatBaptismDateForKey = "0000-00-00"
        Exit Function
    End If

    ' 4. 初始化各部分為 "00" 或 "0000"
    yearPart = "0000"
    monthPart = "00"
    dayPart = "00"

    ' 5. 嘗試解析輸入值
    On Error Resume Next ' 忽略可能的 Split 錯誤

    ' 嘗試按 "-" 分割
    If InStr(rawValue, "-") > 0 Then
        parts = Split(rawValue, "-")
        Select Case UBound(parts)
            Case 2 ' YYYY-MM-DD or similar
                yearPart = ValidateAndPadPart(parts(0), "0000", 4)
                monthPart = ValidateAndPadPart(parts(1), "00", 2)
                dayPart = ValidateAndPadPart(parts(2), "00", 2)
            Case 1 ' YYYY-MM or similar
                yearPart = ValidateAndPadPart(parts(0), "0000", 4)
                monthPart = ValidateAndPadPart(parts(1), "00", 2)
                ' dayPart 保持 "00"
            Case 0 ' YYYY- or similar (只處理第一部分)
                 yearPart = ValidateAndPadPart(parts(0), "0000", 4)
                 ' monthPart 和 dayPart 保持 "00"
            Case Else ' 無法解析，保持預設值
                 ' Do nothing, parts remain default "0000", "00", "00"
        End Select
    ' 嘗試按 "/" 分割 (如果沒有 "-")
    ElseIf InStr(rawValue, "/") > 0 Then
         parts = Split(rawValue, "/")
         ' 假設是 YYYY/MM/DD 或 MM/DD/YYYY - 需要判斷!
         ' 為了簡化，我們優先假設 YYYY/MM/DD (如果第一部分是4位數)
         If UBound(parts) = 2 Then
              If Len(Trim(parts(0))) = 4 And IsNumeric(Trim(parts(0))) Then ' 可能是 YYYY/MM/DD
                  yearPart = ValidateAndPadPart(parts(0), "0000", 4)
                  monthPart = ValidateAndPadPart(parts(1), "00", 2)
                  dayPart = ValidateAndPadPart(parts(2), "00", 2)
              ' Add logic here if you need to handle MM/DD/YYYY or DD/MM/YYYY
              ' ElseIf ... (判斷其他格式)
              Else ' 無法確定格式，保持預設
                   ' Do nothing
              End If
         ElseIf UBound(parts) = 1 Then ' YYYY/MM?
              If Len(Trim(parts(0))) = 4 And IsNumeric(Trim(parts(0))) Then
                  yearPart = ValidateAndPadPart(parts(0), "0000", 4)
                  monthPart = ValidateAndPadPart(parts(1), "00", 2)
                  ' dayPart 保持 "00"
              End If
          ElseIf UBound(parts) = 0 Then ' YYYY/?
               yearPart = ValidateAndPadPart(parts(0), "0000", 4)
          End If
    ' 處理純數字 (可能是年份)
    ElseIf IsNumeric(rawValue) Then
        If Len(rawValue) = 4 Then
            yearPart = ValidateAndPadPart(rawValue, "0000", 4)
            ' monthPart 和 dayPart 保持 "00"
        ' Add handling for other numeric lengths if needed (e.g., YYYYMMDD)
        ' ElseIf Len(rawValue) = 8 Then ...
        Else ' 其他長度數字，無法確定為年份，保持預設
             ' Do nothing
        End If
    ' 處理 Excel 日期類型 (如果前面都沒匹配上)
    ElseIf IsDate(cellValue) Then
         ' 先格式化為標準 YYYY-MM-DD
         Dim formattedExcelDate As String
         formattedExcelDate = Format(cellValue, "yyyy-mm-dd")
         parts = Split(formattedExcelDate, "-")
         If UBound(parts) = 2 Then
             yearPart = ValidateAndPadPart(parts(0), "0000", 4)
             monthPart = ValidateAndPadPart(parts(1), "00", 2)
             dayPart = ValidateAndPadPart(parts(2), "00", 2)
         End If
    ' 其他無法解析的字串，保持預設值 "0000-00-00"
    Else
        ' Do nothing, parts remain default
    End If

    On Error GoTo 0 ' 恢復錯誤處理

    ' 6. 組裝最終結果
    FormatBaptismDateForKey = yearPart & "-" & monthPart & "-" & dayPart

    ' Debug.Print "FormatBaptismDateForKey: Input='" & CStr(cellValue) & "', Output='" & FormatBaptismDateForKey & "'"
End Function

' --- Helper Function: 驗證並填充日期部分 ---
Private Function ValidateAndPadPart(partValue As String, defaultVal As String, numDigits As Integer) As String
    Dim tempVal As String
    Dim numericVal As Double

    tempVal = Trim(partValue)

    ' 如果為空或非數字，返回預設值
    If tempVal = "" Or Not IsNumeric(tempVal) Then
        ValidateAndPadPart = defaultVal
        Exit Function
    End If

    ' 轉換為數字判斷是否為 0
    numericVal = CDbl(tempVal)
    If numericVal = 0 Then
        ValidateAndPadPart = String(numDigits, "0") ' 返回 "00" 或 "0000"
    Else
        ' 對於有效數字，格式化為指定位數（帶前導零）
        ValidateAndPadPart = Format(tempVal, String(numDigits, "0"))
    End If

End Function


Sub TestFormatBaptismDateForKey()

    Dim testValues As Variant
    Dim expectedValues As Variant
    Dim i As Integer
    Dim inputValue As Variant
    Dim actualOutput As String
    Dim result As String
    Dim inputDisplay As String

    ' --- 定義測試案例 ---
    ' 包含各種可能的輸入情況
    testValues = Array( _
        "", _
        Null, _
        "0000-00-00", _
        "0000-08", _
        "2023-12", _
        1997, _
        "2021", _
        " 1985 ", _
        "2021-01-01", _
        "2022/11/05", _
        DateSerial(1999, 3, 15), _
        "1997-01-00", _
        "2021-00-00", _
        "abc", _
        0, _
        "2005-06-00", _
        " 1998-05 ", _
        CVErr(xlErrNA), _
        " - ", _
        "2024-02", _
        2024 _
    )

    ' --- 定義預期輸出 ---
    ' 與 testValues 一一對應
    expectedValues = Array( _
        "0000-00-00", _
        "0000-00-00", _
        "0000-00-00", _
        "1997-01-00", _
        "2023-12-00", _
        "1997-00-00", _
        "2021-00-00", _
        "1985-00-00", _
        "2021-01-01", _
        "2022-11-05", _
        "1999-03-15", _
        "1997-01-00", _
        "2021-00-00", _
        "abc", _
        "0", _
        "2005-06-00", _
        "1998-05-00", _
        "#ERROR!", _
        "-", _
        "2024-02-00", _
        "2024-00-00" _
    )

    ' --- 開始測試並輸出結果到即時運算視窗 (Ctrl+G) ---
    Debug.Print "--- 開始測試 FormatBaptismDateForKey ---"
    Debug.Print String(80, "-") ' 分隔線
    Debug.Print "輸入值", "預期輸出", "實際輸出", "結果"
    Debug.Print String(80, "-")

    For i = LBound(testValues) To UBound(testValues)
        inputValue = testValues(i)

        ' 為了方便顯示，處理特殊輸入值
        If IsNull(inputValue) Then
            inputDisplay = "Null"
        ElseIf IsError(inputValue) Then
            inputDisplay = "#N/A" ' 假設測試 #N/A
        ElseIf IsDate(inputValue) Then
             inputDisplay = CStr(inputValue) & " (Date)" ' 標示為日期類型
        Else
            inputDisplay = "'" & CStr(inputValue) & "'" ' 用引號括起來表示原始輸入
        End If

        ' 調用被測試的函數
        actualOutput = FormatBaptismDateForKey(inputValue)

        ' 比較結果
        If actualOutput = expectedValues(i) Then
            result = "PASS"
        Else
            result = "FAIL"
        End If

        ' 輸出結果
        Debug.Print inputDisplay, "'" & expectedValues(i) & "'", "'" & actualOutput & "'", result
    Next i

    Debug.Print String(80, "-")
    Debug.Print "--- 測試結束 ---"
    MsgBox "測試完成，請查看 VBA 編輯器的即時運算視窗 (Ctrl+G) 獲取結果。", vbInformation

End Sub

