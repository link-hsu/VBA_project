太棒 — 我幫你寫好一個可直接拿來用的 VBA 函數 + 範例 Sub，功能如下：
	•	輸入：兩個檔案路徑（A、B）、要比對的分頁名稱（若有多個請以 ; 分隔）、和要比對的儲存格區段字串（像你給的 "C11:J36,C37:E47,G37:J47,F40,..."，以逗號分隔多個區段）。
	•	行為：逐一比對每個分頁中每個區段的每顆儲存格值。
	•	若完全一致 → 函數回傳數字 0（表示無差異）。
	•	若有不一致 → 函數回傳一個字串，列出所有差異：檔案A 名稱 vs 檔案B 名稱 | 分頁 | 儲存格 | A值 | B值（每個差異一行）。
	•	比對邏輯：
	•	若雙方都是數值，採數值比較（允許非常小的浮點誤差）。
	•	若雙方都是日期，採日期比較。
	•	其它資料型別採字串比對（會 Trim 空白再比）。
	•	它會自動以只讀方式開啟檔案並在完成後關閉。

把下面整段複製到你的 VBA 模組（Module）內即可使用。

Option Explicit

' 比較兩個檔案中相同分頁、相同儲存格區段的所有儲存格
' 回傳 Variant：若無差異回傳 0（數字），否則回傳包含所有差異的字串（每行一筆差異）
' 使用方式範例：
'   res = CompareWorkbooksRanges("C:\A.xlsx", "C:\B.xlsx", "Sheet1;Sheet2", "C11:J36,C37:E47,F40")
Public Function CompareWorkbooksRanges(pathA As String, pathB As String, _
                                      sheetNames As String, ranges As String) As Variant
    Dim wbA As Workbook, wbB As Workbook
    Dim arrSheets() As String, arrRanges() As String
    Dim iS As Long, iR As Long
    Dim wsA As Worksheet, wsB As Worksheet
    Dim rngA As Range, rngB As Range
    Dim r As Long, c As Long
    Dim cellA As Range, cellB As Range
    Dim results As String
    Dim fNameA As String, fNameB As String
    Dim openedA As Boolean, openedB As Boolean
    Dim eps As Double: eps = 0.0000001 ' 數值比較的微小容忍誤差
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 開檔（只讀，不更新連結）
    Set wbA = Workbooks.Open(Filename:=pathA, ReadOnly:=True, UpdateLinks:=0)
    openedA = True
    Set wbB = Workbooks.Open(Filename:=pathB, ReadOnly:=True, UpdateLinks:=0)
    openedB = True
    
    fNameA = wbA.Name
    fNameB = wbB.Name
    
    ' 切分 sheet 名稱（以分號為分隔；若只有一個也可直接傳入）
    arrSheets = Split(sheetNames, ";")
    ' 切分 ranges（以逗號為分隔）
    arrRanges = Split(ranges, ",")
    
    results = ""
    
    For iS = LBound(arrSheets) To UBound(arrSheets)
        Dim sName As String
        sName = Trim(arrSheets(iS))
        If sName = "" Then GoTo NextSheet
        ' 確認工作表存在
        If Not SheetExists(wbA, sName) Then
            results = results & "ERROR: sheet '" & sName & "' not found in " & fNameA & vbCrLf
            GoTo NextSheet
        End If
        If Not SheetExists(wbB, sName) Then
            results = results & "ERROR: sheet '" & sName & "' not found in " & fNameB & vbCrLf
            GoTo NextSheet
        End If
        
        Set wsA = wbA.Worksheets(sName)
        Set wsB = wbB.Worksheets(sName)
        
        For iR = LBound(arrRanges) To UBound(arrRanges)
            Dim token As String
            token = Trim(arrRanges(iR))
            If token = "" Then GoTo NextRange
            
            On Error Resume Next
            Set rngA = wsA.Range(token)
            Set rngB = wsB.Range(token)
            If rngA Is Nothing Or rngB Is Nothing Then
                results = results & "ERROR: 範圍 '" & token & "' 在分頁 '" & sName & "' 無法解析" & vbCrLf
                On Error GoTo ErrHandler
                GoTo NextRange
            End If
            On Error GoTo ErrHandler
            
            ' 形狀應相同（同列數同欄數）
            If rngA.Rows.Count <> rngB.Rows.Count Or rngA.Columns.Count <> rngB.Columns.Count Then
                results = results & "ERROR: 範圍 '" & token & "' 在分頁 '" & sName & "' 大小不相同" & vbCrLf
                GoTo NextRange
            End If
            
            ' 逐格比對（以左上為基準）
            For r = 1 To rngA.Rows.Count
                For c = 1 To rngA.Columns.Count
                    Set cellA = rngA.Cells(r, c)
                    Set cellB = rngB.Cells(r, c)
                    If Not CellsEqual(cellA.Value, cellB.Value, eps) Then
                        results = results & fNameA & " vs " & fNameB & " | " & _
                                  "Sheet: " & sName & " | Cell: " & cellA.Address(False, False) & _
                                  " | A: " & SafeToString(cellA.Value) & " | B: " & SafeToString(cellB.Value) & vbCrLf
                    End If
                Next c
            Next r
            
NextRange:
            Set rngA = Nothing
            Set rngB = Nothing
        Next iR
NextSheet:
    Next iS
    
    ' 回傳：若 results 為空 -> 回傳數字 0；否則回傳結果字串
    If results = "" Then
        CompareWorkbooksRanges = 0
    Else
        CompareWorkbooksRanges = results
    End If

Cleanup:
    On Error Resume Next
    If openedA Then wbA.Close SaveChanges:=False
    If openedB Then wbB.Close SaveChanges:=False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Function

ErrHandler:
    results = results & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf
    Resume Cleanup
End Function

' 檢查工作表是否存在
Private Function SheetExists(wb As Workbook, sName As String) As Boolean
    On Error Resume Next
    SheetExists = Not wb.Worksheets(sName) Is Nothing
    On Error GoTo 0
End Function

' 安全轉字串（處理 Null / Empty / Error）
Private Function SafeToString(v As Variant) As String
    If IsError(v) Then
        SafeToString = "#ERR"
    ElseIf IsNull(v) Then
        SafeToString = "Null"
    ElseIf IsEmpty(v) Then
        SafeToString = ""
    Else
        SafeToString = CStr(v)
    End If
End Function

' 比較兩個儲存格的 Value（支援數值、日期、字串）
' eps 為數值比較的容忍誤差
Private Function CellsEqual(a As Variant, b As Variant, Optional eps As Double = 0.0000001) As Boolean
    On Error GoTo CErr
    ' Error 值視為不相等（但在 SafeToString 會標示）
    If IsError(a) Or IsError(b) Then
        CellsEqual = False
        Exit Function
    End If
    ' Null / Empty 處理
    If (IsNull(a) Or IsEmpty(a)) And (IsNull(b) Or IsEmpty(b)) Then
        CellsEqual = True
        Exit Function
    End If
    ' 兩邊都是數字
    If IsNumeric(a) And IsNumeric(b) Then
        CellsEqual = (Abs(CDbl(a) - CDbl(b)) <= eps)
        Exit Function
    End If
    ' 兩邊都是日期
    If IsDate(a) And IsDate(b) Then
        CellsEqual = (CDate(a) = CDate(b))
        Exit Function
    End If
    ' 其他以字串比對（Trim）
    CellsEqual = (Trim(CStr(a)) = Trim(CStr(b)))
    Exit Function
CErr:
    CellsEqual = False
End Function

' 範例 Sub：如何呼叫並把結果輸出到新工作表
Public Sub Example_CompareAndOutput()
    Dim pathA As String, pathB As String
    Dim sheets As String, ranges As String
    Dim res As Variant
    pathA = "C:\temp\A.xlsx"
    pathB = "C:\temp\B.xlsx"
    sheets = "Sheet1" ' 若要多個 sheet，用分號 ; 分隔，例如 "Sheet1;Sheet2"
    ranges = "C11:J36,C37:E47,G37:J47,F40,F42:F47,C48:C50,J48:J50,C52:C56,J52:J56,C57:J57"
    
    res = CompareWorkbooksRanges(pathA, pathB, sheets, ranges)
    
    If VarType(res) = vbDouble Or VarType(res) = vbInteger Then
        MsgBox "比對結果：無差異（回傳 0）", vbInformation
    Else
        ' 若差異很多，建議輸出到新工作表
        Dim w As Workbook, outWS As Worksheet, lines() As String, i As Long
        Set w = ThisWorkbook
        Set outWS = w.Sheets.Add(After:=w.Sheets(w.Sheets.Count))
        outWS.Name = "CompareResults_" & Format(Now, "hhmmss")
        lines = Split(res, vbCrLf)
        For i = 0 To UBound(lines)
            If Trim(lines(i)) <> "" Then outWS.Cells(i + 1, 1).Value = lines(i)
        Next i
        MsgBox "比對完成，發現差異。結果已寫入工作表：" & outWS.Name, vbExclamation
    End If
End Sub


⸻

如果你希望我做其中一件事，我可以幫你：
	•	把結果改成回傳 Collection 或 Dictionary（程式內更好處理）。
	•	加入「忽略空白／格式差異只比數值」的選項。
	•	支援以「通配字元的 sheet 名稱」或在不同 sheet 對應不同 ranges（目前是同一組 ranges 套用到每個指定 sheet）。
	•	或把差異輸出成 CSV / 新檔案或更漂亮的報表格式。

要我直接幫你調整哪個行為？還是你要我把這段改成 Sub（不回傳，而是直接產生報表）？




' ========
' Question

以下地方請幫我修改
1.傳入的sheets要以Array傳入，不要以逗號分隔
2.最後回傳結果要回傳的是結構化的資料，方便我後續處理，
如果比對沒有問題則回傳0或Null這種可以從if判斷式立刻檢核出的結果。

' ======
' Answer


好，已依你的兩個要求改好了：
	•	sheets 改成以 Array（或任何 Variant array）傳入，不再用逗號分隔字串。呼叫範例會示範用 Array("Sheet1","Sheet2") 傳入。
	•	回傳改為結構化資料：若沒有差異回傳數字 0（方便直接在 If 判斷式使用）；若有差異回傳一個 Collection，Collection 裡每一項都是一個 Scripting.Dictionary（用 late binding 建立，不需額外 Reference），字典欄位如下：
	•	"Type": "Diff" 或 "Error"
	•	"FileA" / "FileB": 檔名
	•	"Sheet": 分頁名稱
	•	"Cell": 儲存格位置（A1 格式）
	•	"ValueA" / "ValueB": 兩邊的值（若 Type="Error"，則 "ValueA" 放錯誤訊息）

下面把完整的 VBA 程式給你，直接貼到 Module 中就能用：

Option Explicit

' 比較兩個檔案中相同分頁、相同儲存格區段的所有儲存格
' 參數：
'   pathA, pathB: 完整檔案路徑
'   sheetNames: 傳入 Array，例如 Array("Sheet1","Sheet2") 或 Variant array
'   ranges: 以逗號分隔的區段字串，例如 "C11:J36,C37:E47,F40"
' 回傳：
'   若無差異 -> 回傳數字 0
'   若有差異或錯誤 -> 回傳 Collection (每項為 Scripting.Dictionary)
Public Function CompareWorkbooksRanges(pathA As String, pathB As String, _
                                      sheetNames As Variant, ranges As String) As Variant
    Dim wbA As Workbook, wbB As Workbook
    Dim arrSheets As Variant, arrRanges() As String
    Dim iS As Long, iR As Long
    Dim wsA As Worksheet, wsB As Worksheet
    Dim rngA As Range, rngB As Range
    Dim r As Long, c As Long
    Dim cellA As Range, cellB As Range
    Dim resultCol As Collection
    Dim dict As Object ' Scripting.Dictionary (late binding)
    Dim fNameA As String, fNameB As String
    Dim openedA As Boolean, openedB As Boolean
    Dim eps As Double: eps = 0.0000001 ' 數值比較的容忍誤差
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 檢查並建立 sheet array
    If IsArray(sheetNames) Then
        arrSheets = sheetNames
    Else
        ' 若使用者誤傳單一字串，也處理：轉成一元素的 array
        ReDim arrSheets(0 To 0)
        arrSheets(0) = CStr(sheetNames)
    End If
    
    arrRanges = Split(ranges, ",")
    Set resultCol = New Collection
    
    ' 開檔（只讀，不更新連結）
    Set wbA = Workbooks.Open(Filename:=pathA, ReadOnly:=True, UpdateLinks:=0)
    openedA = True
    Set wbB = Workbooks.Open(Filename:=pathB, ReadOnly:=True, UpdateLinks:=0)
    openedB = True
    
    fNameA = wbA.Name
    fNameB = wbB.Name
    
    For iS = LBound(arrSheets) To UBound(arrSheets)
        Dim sName As String
        sName = Trim(CStr(arrSheets(iS)))
        If sName = "" Then GoTo NextSheetLoop
        
        ' 檢查工作表存在性
        If Not SheetExists(wbA, sName) Then
            Set dict = CreateObject("Scripting.Dictionary")
            dict("Type") = "Error"
            dict("FileA") = fNameA
            dict("FileB") = fNameB
            dict("Sheet") = sName
            dict("Cell") = ""
            dict("ValueA") = "Sheet not found in " & fNameA
            dict("ValueB") = ""
            resultCol.Add dict
            GoTo NextSheetLoop
        End If
        If Not SheetExists(wbB, sName) Then
            Set dict = CreateObject("Scripting.Dictionary")
            dict("Type") = "Error"
            dict("FileA") = fNameA
            dict("FileB") = fNameB
            dict("Sheet") = sName
            dict("Cell") = ""
            dict("ValueA") = ""
            dict("ValueB") = "Sheet not found in " & fNameB
            resultCol.Add dict
            GoTo NextSheetLoop
        End If
        
        Set wsA = wbA.Worksheets(sName)
        Set wsB = wbB.Worksheets(sName)
        
        For iR = LBound(arrRanges) To UBound(arrRanges)
            Dim token As String
            token = Trim(arrRanges(iR))
            If token = "" Then GoTo NextRangeLoop
            
            On Error Resume Next
            Set rngA = wsA.Range(token)
            Set rngB = wsB.Range(token)
            If rngA Is Nothing Or rngB Is Nothing Then
                Set dict = CreateObject("Scripting.Dictionary")
                dict("Type") = "Error"
                dict("FileA") = fNameA
                dict("FileB") = fNameB
                dict("Sheet") = sName
                dict("Cell") = token
                dict("ValueA") = IIf(rngA Is Nothing, "Range not found/invalid in A", "")
                dict("ValueB") = IIf(rngB Is Nothing, "Range not found/invalid in B", "")
                resultCol.Add dict
                On Error GoTo ErrHandler
                GoTo NextRangeLoop
            End If
            On Error GoTo ErrHandler
            
            ' 檢查大小一致
            If rngA.Rows.Count <> rngB.Rows.Count Or rngA.Columns.Count <> rngB.Columns.Count Then
                Set dict = CreateObject("Scripting.Dictionary")
                dict("Type") = "Error"
                dict("FileA") = fNameA
                dict("FileB") = fNameB
                dict("Sheet") = sName
                dict("Cell") = token
                dict("ValueA") = "Range size: " & rngA.Rows.Count & "x" & rngA.Columns.Count
                dict("ValueB") = "Range size: " & rngB.Rows.Count & "x" & rngB.Columns.Count
                resultCol.Add dict
                GoTo NextRangeLoop
            End If
            
            ' 逐格比對
            For r = 1 To rngA.Rows.Count
                For c = 1 To rngA.Columns.Count
                    Set cellA = rngA.Cells(r, c)
                    Set cellB = rngB.Cells(r, c)
                    If Not CellsEqual(cellA.Value, cellB.Value, eps) Then
                        Set dict = CreateObject("Scripting.Dictionary")
                        dict("Type") = "Diff"
                        dict("FileA") = fNameA
                        dict("FileB") = fNameB
                        dict("Sheet") = sName
                        dict("Cell") = cellA.Address(False, False)
                        dict("ValueA") = SafeToString(cellA.Value)
                        dict("ValueB") = SafeToString(cellB.Value)
                        resultCol.Add dict
                    End If
                Next c
            Next r
            
NextRangeLoop:
            Set rngA = Nothing
            Set rngB = Nothing
        Next iR
NextSheetLoop:
    Next iS
    
    ' 若沒有任何差異或錯誤則回傳 0，否則回傳 Collection
    If resultCol.Count = 0 Then
        CompareWorkbooksRanges = 0
    Else
        Set CompareWorkbooksRanges = resultCol
    End If

Cleanup:
    On Error Resume Next
    If openedA Then wbA.Close SaveChanges:=False
    If openedB Then wbB.Close SaveChanges:=False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Function

ErrHandler:
    ' 如果發生未預期錯誤，回傳一個包含該錯誤的 Dictionary 於 Collection 中
    On Error Resume Next
    Set dict = CreateObject("Scripting.Dictionary")
    dict("Type") = "Error"
    dict("FileA") = IIf(openedA, fNameA, pathA)
    dict("FileB") = IIf(openedB, fNameB, pathB)
    dict("Sheet") = ""
    dict("Cell") = ""
    dict("ValueA") = "Runtime Error " & Err.Number & ": " & Err.Description
    dict("ValueB") = ""
    If resultCol Is Nothing Then Set resultCol = New Collection
    resultCol.Add dict
    Resume Cleanup
End Function

' 檢查工作表是否存在
Private Function SheetExists(wb As Workbook, sName As String) As Boolean
    On Error Resume Next
    SheetExists = Not wb.Worksheets(sName) Is Nothing
    On Error GoTo 0
End Function

' 安全轉字串（處理 Error / Null / Empty）
Private Function SafeToString(v As Variant) As String
    If IsError(v) Then
        SafeToString = "#ERR"
    ElseIf IsNull(v) Then
        SafeToString = "Null"
    ElseIf IsEmpty(v) Then
        SafeToString = ""
    Else
        SafeToString = CStr(v)
    End If
End Function

' 比較兩個儲存格的 Value（支援數值、日期、字串）
Private Function CellsEqual(a As Variant, b As Variant, Optional eps As Double = 0.0000001) As Boolean
    On Error GoTo CErr
    If IsError(a) Or IsError(b) Then
        CellsEqual = False
        Exit Function
    End If
    If (IsNull(a) Or IsEmpty(a)) And (IsNull(b) Or IsEmpty(b)) Then
        CellsEqual = True
        Exit Function
    End If
    If IsNumeric(a) And IsNumeric(b) Then
        CellsEqual = (Abs(CDbl(a) - CDbl(b)) <= eps)
        Exit Function
    End If
    If IsDate(a) And IsDate(b) Then
        CellsEqual = (CDate(a) = CDate(b))
        Exit Function
    End If
    CellsEqual = (Trim(CStr(a)) = Trim(CStr(b)))
    Exit Function
CErr:
    CellsEqual = False
End Function


⸻

使用範例（範例 Sub，示範如何傳入 Array、以及如何判斷回傳結果並處理 Collection）：

Public Sub Example_UseCompare()
    Dim pathA As String, pathB As String
    Dim sheetsArr As Variant
    Dim ranges As String
    Dim res As Variant
    Dim diffs As Collection
    Dim i As Long
    pathA = "C:\temp\A.xlsx"
    pathB = "C:\temp\B.xlsx"
    sheetsArr = Array("Sheet1") ' 或 Array("Sheet1","Sheet2")
    ranges = "C11:J36,C37:E47,G37:J47,F40"
    
    res = CompareWorkbooksRanges(pathA, pathB, sheetsArr, ranges)
    
    If VarType(res) = vbDouble And res = 0 Then
        MsgBox "比對結果：無差異", vbInformation
        Exit Sub
    End If
    
    ' 若回傳為 Collection（有差異或錯誤），逐筆處理
    If IsObject(res) Then
        Set diffs = res
        For i = 1 To diffs.Count
            Dim d As Object
            Set d = diffs(i)
            If d("Type") = "Diff" Then
                Debug.Print "DIFF - Sheet:" & d("Sheet") & " Cell:" & d("Cell") & _
                            " A=" & d("ValueA") & " B=" & d("ValueB")
            Else
                Debug.Print "ERROR - " & d("ValueA") & " " & d("ValueB")
            End If
        Next i
    Else
        MsgBox "未預期的回傳型別", vbExclamation
    End If
End Sub


⸻

需要我再幫你做的事（選一個或多個）：
	•	把 Collection 中 Dictionary 的欄位改成你偏好的欄位名稱或額外欄位（例如加入 Range 原始 token）。
	•	改成回傳 Dictionary（key 為 Sheet|Cell）而非 Collection。
	•	將比對結果自動寫成 Excel 報表（CSV / 新工作表）而不是只回傳結構化物件。

要我立刻把哪一項改成你想要的形式？




' ============
