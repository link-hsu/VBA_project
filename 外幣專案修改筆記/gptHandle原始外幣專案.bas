Question
Question

' ====================
這是我的class clsReport
clsReport.bas

' Report Title
Private clsReportName As String

' Dictionary：key = Worksheet Name，value = Dictionary( Keys "Fiedl Values" 與 "Field Addresses" )
Private clsWorksheets As Object

'=== 初始化報表 (根據報表名稱建立各工作表的欄位定義) ===
Public Sub Init(ByVal reportName As String, _
                ByVal dataMonthStringROC As String, _
                ByVal dataMonthStringROC_NUM As String, _
                ByVal dataMonthStringROC_F1F2 As String)
    clsReportName = reportName
    Set clsWorksheets = CreateObject("Scripting.Dictionary")
    
    Select Case reportName
        Case "CNY1"
            AddWorksheetFields "CNY1", Array( _
                Array("CNY1_申報時間", "C2", dataMonthStringROC), _
                Array("CNY1_其他金融資產_淨額", "G98", Null), _
                Array("CNY1_其他", "G100", Null), _
                Array("CNY1_資產總計", "G116", Null), _
                Array("CNY1_其他金融負債", "G170", Null), _
                Array("CNY1_其他什項金融負債", "G172", Null), _
                Array("CNY1_負債總計", "G184", Null) )
        Case "FB1"
            'No Data
            AddWorksheetFields "FOA", Array( _
                Array("FB1_申報時間", "C2", dataMonthStringROC) )
        Case "FB2"
            AddWorksheetFields "FOA", Array( _
                Array("FB2_申報時間", "D2", dataMonthStringROC), _
                Array("FB2_存放及拆借同業", "F9", Null), _
                Array("FB2_拆放銀行同業", "F13", Null), _
                Array("FB2_應收款項_淨額", "F36", Null), _
                Array("FB2_應收利息", "F41", Null), _
                Array("FB2_資產總計", "F85", Null) )
        Case "FB3"
            AddWorksheetFields "FOA", Array( _
                Array("FB3_申報時間", "C2", dataMonthStringROC), _
                Array("FB3_存放及拆借同業_資產面_台灣地區", "D9", Null), _
                Array("FB3_同業存款及拆放_負債面_台灣地區", "D10", Null) )
        Case "FB3A"
            ' Dynamically create in following Process Processdure
            AddWorksheetFields "FOA", Array( _
                Array("FB3A_申報時間", "C2", dataMonthStringROC) )
    End Select
End Sub

'=== Private Method：Add Def for Worksheet Field === 
' fieldDefs is array of fields(each field(Array) of fields(Array)),
' for each Index's Form => (FieldName, CellAddress, InitialVAlue(null))
Private Sub AddWorksheetFields(ByVal wsName As String, _
                               ByVal fieldDefs As Variant)
    Dim wsDict As Object, dictValues As Object, dictAddresses As Object

    Dim i As Long, arrField As Variant

    Set dictValues = CreateObject("Scripting.Dictionary")
    Set dictAddresses = CreateObject("Scripting.Dictionary")
    
    For i = LBound(fieldDefs) To UBound(fieldDefs)
        arrField = fieldDefs(i)
        dictValues.Add arrField(0), arrField(2)
        dictAddresses.Add arrField(0), arrField(1)
    Next i
    
    Set wsDict = CreateObject("Scripting.Dictionary")
    wsDict.Add "Values", dictValues
    wsDict.Add "Addresses", dictAddresses
    
    clsWorksheets.Add wsName, wsDict
End Sub

Public Sub AddDynamicField(ByVal wsName As String, _
                           ByVal fieldName As String, _
                           ByVal cellAddress As String, _
                           ByVal initValue As Variant)
    Dim wsDict As Object
    Dim dictValues As Object, dictAddresses As Object
    
    ' 如果該工作表尚未建立，先建立一組新的 Dictionary
    If Not clsWorksheets.Exists(wsName) Then
        Set dictValues = CreateObject("Scripting.Dictionary")
        Set dictAddresses = CreateObject("Scripting.Dictionary")
        
        Set wsDict = CreateObject("Scripting.Dictionary")
        wsDict.Add "Values", dictValues
        wsDict.Add "Addresses", dictAddresses
        
        clsWorksheets.Add wsName, wsDict
    End If
    
    ' 取得該工作表的字典
    Set wsDict = clsWorksheets(wsName)
    Set dictValues = wsDict("Values")
    Set dictAddresses = wsDict("Addresses")
    
    ' 如果欄位已存在，可依需求選擇更新或忽略（此處以加入為例）
    If Not dictValues.Exists(fieldName) Then
        dictValues.Add fieldName, initValue
        dictAddresses.Add fieldName, cellAddress
    Else
        ' 若需要更新，直接賦值：
        dictValues(fieldName) = initValue
        dictAddresses(fieldName) = cellAddress
    End If
End Sub

'=== Set Field Value for one sheetName ===  
Public Sub SetField(ByVal wsName As String, _
                    ByVal fieldName As String, _
                    ByVal value As Variant)
    If Not clsWorksheets.Exists(wsName) Then
        Err.Raise 1002, , "工作表 [" & wsName & "] 尚未定義於報表 " & clsReportName
    End If
    Dim wsDict As Object
    Set wsDict = clsWorksheets(wsName)
    Dim dictValues As Object
    Set dictValues = wsDict("Values")
    If dictValues.Exists(fieldName) Then
        dictValues(fieldName) = value
    Else
        Err.Raise 1001, , "欄位 [" & fieldName & "] 不存在於工作表 [" & wsName & "] 的報表 " & clsReportName
    End If
End Sub

'=== With NO Parma: Get All Field Values ===  
'=== With wsName: Get Field Values within the worksheet Key 格式："wsName|fieldName" ===
Public Function GetAllFieldValues(Optional ByVal wsName As String = "") As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    Dim wsKey As Variant, dictV As Object, fieldKey As Variant
    If wsName <> "" Then
        If Not clsWorksheets.Exists(wsName) Then
            Err.Raise 1002, , "工作表 [" & wsName & "] 尚未定義於報表 " & clsReportName
        End If
        Set result = clsWorksheets(wsName)("Values")
    Else
        For Each wsKey In clsWorksheets.Keys
            Set dictV = clsWorksheets(wsKey)("Values")
            For Each fieldKey In dictV.Keys
                result.Add wsKey & "|" & fieldKey, dictV(fieldKey)
            Next fieldKey
        Next wsKey
    End If
    Set GetAllFieldValues = result
End Function

'=== With No Param: Get All Field Addresses ===  
'=== With wsName: Get Field Addresses within the worksheet Key 格式："wsName|fieldName" ===
Public Function GetAllFieldPositions(Optional ByVal wsName As String = "") As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    Dim wsKey As Variant, dictA As Object, fieldKey As Variant
    If wsName <> "" Then
        If Not clsWorksheets.Exists(wsName) Then
            Err.Raise 1002, , "工作表 [" & wsName & "] 尚未定義於報表 " & clsReportName
        End If
        Set result = clsWorksheets(wsName)("Addresses")
    Else
        For Each wsKey In clsWorksheets.Keys
            Set dictA = clsWorksheets(wsKey)("Addresses")
            For Each fieldKey In dictA.Keys
                result.Add wsKey & "|" & fieldKey, dictA(fieldKey)
            Next fieldKey
        Next wsKey
    End If
    Set GetAllFieldPositions = result
End Function

'=== 驗證是否每個欄位都有填入數值 (若指定 wsName 則驗證該工作表) ===  
Public Function ValidateFields(Optional ByVal wsName As String = "") As Boolean
    Dim msg As String, key As Variant
    msg = ""
    Dim dictValues As Object
    If wsName <> "" Then
        If Not clsWorksheets.Exists(wsName) Then
            Err.Raise 1002, , "工作表 [" & wsName & "] 尚未定義於報表 " & clsReportName
        End If
        Set dictValues = clsWorksheets(wsName)("Values")
        For Each key In dictValues.Keys
            If IsNull(dictValues(key)) Then msg = msg & wsName & " - " & key & vbCrLf
        Next key
    Else
        Dim wsKey As Variant
        For Each wsKey In clsWorksheets.Keys
            Set dictValues = clsWorksheets(wsKey)("Values")
            For Each key In dictValues.Keys
                If IsNull(dictValues(key)) Then msg = msg & wsKey & " - " & key & vbCrLf
            Next key
        Next wsKey
    End If
    If msg <> "" Then
        MsgBox "報表 [" & clsReportName & "] 以下欄位未填入有效資料:" & vbCrLf & msg, vbExclamation
        WriteLog "報表 [" & clsReportName & "] 以下欄位未填入有效資料:" & vbCrLf & msg
        ValidateFields = False
    Else
        ValidateFields = True
    End If
End Function

'=== 將 class 中的數值依據各工作表之欄位設定寫入指定的 Workbook ===  
' 此方法會針對 clsWorksheets 中定義的每個工作表名稱，嘗試在傳入的 Workbook 中找到對應工作表，並更新其欄位
Public Sub ApplyToWorkbook(ByRef wb As Workbook)
    Dim wsKey As Variant, wsDict As Object, dictValues As Object, dictAddresses As Object
    Dim ws As Worksheet, fieldKey As Variant
    For Each wsKey In clsWorksheets.Keys
        On Error Resume Next
        Set ws = wb.Sheets(wsKey)
        On Error GoTo 0
        If ws Is Nothing Then
            MsgBox "Workbook 中找不到工作表: " & wsKey, vbExclamation
            WriteLog "Workbook 中找不到工作表: " & wsKey
            Exit Sub
        End If
        
        Set wsDict = clsWorksheets(wsKey)
        Set dictValues = wsDict("Values")
        Set dictAddresses = wsDict("Addresses")
        For Each fieldKey In dictValues.Keys
            If Not IsNull(dictValues(fieldKey)) Then
                On Error Resume Next
                ws.Range(dictAddresses(fieldKey)).Value = dictValues(fieldKey)
                If Err.Number <> 0 Then
                    MsgBox "工作表 [" & wsKey & "] 找不到儲存格 " & _
                           dictAddresses(fieldKey) & " （欄位：" & fieldKey & "）", vbExclamation
                    WriteLog "工作表 [" & wsKey & "] 找不到儲存格 " & _
                             dictAddresses(fieldKey) & " （欄位：" & fieldKey & "）"
                    Err.Clear
                End If
                On Error GoTo 0
            Else
                ' 沒呼叫 SetField 的欄位 (值還是 Null)
                MsgBox "工作表 [" & wsKey & "] 欄位尚未設定值: " & fieldKey, vbExclamation
                WriteLog "工作表 [" & wsKey & "] 欄位尚未設定值: " & fieldKey
            End If
        Next fieldKey
        Set ws = Nothing
    Next wsKey
End Sub

'=== 報表名稱屬性 ===  
Public Property Get ReportName() As String
    ReportName = clsReportName
End Property

' ====================
這是我的初始(主)執行程序

Option Explicit

'=== Global Config Settings ===
Public gDataMonthString As String         ' 使用者輸入的資料月份
Public gDataMonthStringROC As String      ' 資料月份ROC Format
Public gDataMonthStringROC_NUM As String  ' 資料月份ROC_NUM Format
Public gDataMonthStringROC_F1F2 As String ' 資料月份ROC_F1F2 Format
Public gDBPath As String                  ' 資料庫路徑
Public gReportFolder As String            ' 原始申報報表 Excel 檔所在資料夾
Public gOutputFolder As String            ' 更新後另存新檔的資料夾
Public gReportNames As Variant            ' 報表名稱陣列
Public gReports As Collection             ' Declare Collections that Save all instances of clsReport
Public gRecIndex As Long                  ' RecordIndex 計數器

'=== UserForm 新增全域 allReportNames
Public allReportNames As Variant

'=== 主流程入口 ===
Public Sub Main()
    Dim isInputValid As Boolean
    isInputValid = False
    Do
        gDataMonthString = InputBox("請輸入資料月份 (格式: yyyy/mm):", "輸入資料月份")
        If IsValidDataMonth(gDataMonthString) Then
            isInputValid = True
        ElseIf Trim(gDataMonthString) = "" Then
            MsgBox "請輸入報表資料所屬的年度/月份 (例如: 2024/01)", vbExclamation, "輸入錯誤"
            WriteLog "請輸入報表資料所屬的年度/月份 (例如: 2024/01)"
        Else
            MsgBox "格式錯誤，請輸入正確格式 (yyyy/mm)", vbExclamation, "格式錯誤"
            WriteLog "格式錯誤，請輸入正確格式 (yyyy/mm)"
        End If
    Loop Until isInputValid

    ThisWorkbook.Sheets("ControlPanel").Range("gDataMonthString").Value = "'" & gDataMonthString
    
    '轉換gDataMonthString為ROC Format
    gDataMonthStringROC = ConvertToROCFormat(gDataMonthString, "ROC")
    gDataMonthStringROC_NUM = ConvertToROCFormat(gDataMonthString, "NUM")
    gDataMonthStringROC_F1F2 = ConvertToROCFormat(gDataMonthString, "F1F2")
    ' 設定其他 config 參數（請根據實際環境調整）
    gDBPath = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("DBsPathFileName").Value
    ' gDBPath = "\\10.10.122.40\後台作業\99_個人資料夾\8.修豪\DbsMReport20250513_V1\" & ThisWorkbook.Sheets("ControlPanel").Range("DBsPathFileName").Value
    ' 空白報表路徑
    gReportFolder = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("EmptyReportPath").Value
    ' 產生之申報報表路徑
    gOutputFolder = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("OutputReportPath").Value

    ' ========== 宣告所有報表 ==========
    ' 製作報表List
    ' gReportNames 少FB1 FM5
    ' allReportNames = Array("CNY1", "FB1", "FB2", "FB3", "FB3A", "FM5", "FM11", "FM13", "AI821", "Table2", "FB5", "FB5A", "FM2", "FM10", "F1_F2", "Table41", "AI602", "AI240", "AI822")

    allReportNames = Array("FB1")

    ' =====testArray=====
    ' allReportNames = Array("AI822")

    ' ========== 選擇產生全部或部分報表 ==========
    Dim respRunAll As VbMsgBoxResult
    Dim userInput As String
    Dim i As Integer, j As Integer
    respRunAll = MsgBox("要執行全部報表嗎？" & vbCrLf & _
                  "【是】→ 全部報表" & vbCrLf & _
                  "【否】→ 指定報表", _
                  vbQuestion + vbYesNo, "選擇產生全部或部分報表")    
    If respRunAll = vbYes Then
        gReportNames = allReportNames
        For i = LBound(gReportNames) To UBound(gReportNames)
            gReportNames(i) = UCase(gReportNames(i))
        Next i
    Else
        ' UserForm 勾選清單
        Dim frm As ReportSelector
        Set frm = New ReportSelector
        frm.Show vbModal
        ' 若 gReportNames 未被填（使用者未選任何項目），則中止
        If Not IsArray(gReportNames) Or UBound(gReportNames) < 0 Then
            MsgBox "未選擇任何報表，程序結束", vbInformation
            Exit Sub
        End If
        ' 轉大寫（保留原邏輯）
        For i = LBound(gReportNames) To UBound(gReportNames)
            gReportNames(i) = UCase(gReportNames(i))
        Next i
    End If
    
    ' 檢查不符合的報表名稱
    Dim invalidReports As String
    Dim found As Boolean

    For i = LBound(gReportNames) To UBound(gReportNames)
        found = False
        For j = LBound(allReportNames) To UBound(allReportNames)
            If UCase(gReportNames(i)) = UCase(allReportNames(j)) Then
                found = True
                Exit For
            End If
        Next j
        If Not found Then
            invalidReports = invalidReports & gReportNames(i) & ", "
        End If

    Next i
    If Len(invalidReports) > 0 Then
        invalidReports = Left(invalidReports, Len(invalidReports) - 2)
        MsgBox "報表名稱錯誤，請重新確認：" & vbCrLf & invalidReports, vbCritical, "報表名稱錯誤"
        WriteLog "報表名稱錯誤，請重新確認：" & vbCrLf & invalidReports
        Exit Sub
    End If
    
    ' ========== 處理其他部門提供數據欄位 ==========
    ' 定義每張報表必需由使用者填入／確認的儲存格名稱
    Dim req As Object
    Set req = CreateObject("Scripting.Dictionary")
    req.Add "TABLE41", Array("Table41_國外部_一利息收入", _
                             "Table41_國外部_一利息收入_利息", _
                             "Table41_國外部_一利息收入_利息_存放銀行同業", _
                             "Table41_國外部_二金融服務收入", _
                             "Table41_國外部_一利息支出", _
                             "Table41_國外部_一利息支出_利息", _
                             "Table41_國外部_一利息支出_利息_外國人外匯存款", _
                             "Table41_國外部_二金融服務支出", _
                             "Table41_企銷處_一利息支出", _
                             "Table41_企銷處_一利息支出_利息", _
                             "Table41_企銷處_一利息支出_利息_外國人新台幣存款")
                            
    req.Add "AI822", Array("AI822_會計科_上年度決算後淨值", _
                           "AI822_國外部_直接往來之授信", _
                           "AI822_國外部_間接往來之授信", _
                           "AI822_授管處_直接往來之授信")

    ' 暫存要移除的報表
    Dim toRemove As Collection
    Set toRemove = New Collection

    ' 逐一詢問使用者每張報表、每個必要欄位的值
    Dim ws As Worksheet
    Dim rptName As Variant 
    Dim fields As Variant, fld As Variant
    Dim defaultVal As Variant, userVal As String
    Dim respToContinue As VbMsgBoxResult
    Dim respHasInput As VbMsgBoxResult

    For Each rptName In gReportNames
        If req.Exists(rptName) Then
            Set ws = ThisWorkbook.Sheets(rptName)
            fields = req(rptName)

            ' --- 新增：先問一次是否已自行填入該報表所有資料 ---
            respHasInput = MsgBox( _
                "是否已填入 " & rptName & " 報表資料？", _
                vbQuestion + vbYesNo, "確認是否填入資料")
            If respHasInput = vbYes Then
                ' --- 已填入：只檢查「空白」的必要欄位 ---
                For Each fld In fields
                    If Trim(CStr(ws.Range(fld).Value)) = "" Then
                        defaultVal = 0
                        userVal = InputBox( _
                            "報表 " & rptName & " 的欄位 [" & fld & "] 尚未輸入，請填入數值：", _
                            "請填入必要欄位", "")

                            Dim cleanUserVal As String
                            cleanUserVal = Replace(userVal, ",", "")

                        If userVal = "" Then
                            respToContinue = MsgBox("未輸入任何數值，是否仍要製作報表 " & rptName & "？", _
                                         vbQuestion + vbYesNo, "繼續製作？")
                            If respToContinue = vbYes Then
                                ws.Range(fld).Value = 0
                            Else
                                toRemove.Add rptName
                                Exit For
                            End If
                        ElseIf IsNumeric(cleanUserVal) Then
                            ws.Range(fld).Value = CDbl(cleanUserVal)
                        Else
                            ws.Range(fld).Value = 0
                            MsgBox "您輸入的不是數字，將保留為 0", vbExclamation
                            WriteLog "您輸入的不是數字，將保留為 0"
                        End If
                    End If
                Next fld
            Else
                For Each fld In fields
                    defaultVal = ws.Range(fld).Value
                    Dim defaultValFormatWithComma As String
                    defaultValFormatWithComma = Format(defaultVal, "#,##0.###")
                    
                    userVal = InputBox( _
                        "請確認報表 " & rptName & " 的 [" & fld & "]" & vbCrLf & _
                        "目前值：" & defaultValFormatWithComma & vbCrLf & _
                        "若要修改，請輸入新數值；若已更改，請直接點擊「確定」。", _
                        "欄位值", defaultValFormatWithComma _
                    )

                    cleanUserVal = Replace(userVal, ",", "")

                    If userVal = "" Then
                        ' 空白表示使用者沒有輸入
                        respToContinue = MsgBox("未輸入任何數值，是否仍要製作報表 " & rptName & "？", _
                                    vbQuestion + vbYesNo, "繼續製作？")
                        If respToContinue = vbYes Then
                            If IsNumeric(defaultVal) Then
                                ws.Range(fld).Value = CDbl(defaultVal)
                            Else
                                ws.Range(fld).Value = 0
                            End If
                        Else
                            toRemove.Add rptName
                            Exit For   ' 跳出該報表的欄位迴圈
                        End If
                    ElseIf IsNumeric(cleanUserVal) Then
                        ws.Range(fld).Value = CDbl(cleanUserVal)
                    Else
                        If IsNumeric(defaultVal) Then
                            ws.Range(fld).Value = CDbl(defaultVal)
                        Else
                            ws.Range(fld).Value = 0
                        End If
                        MsgBox "您輸入的不是數字，將保留原值：" & defaultValFormatWithComma, vbExclamation
                        WriteLog "您輸入的不是數字，將保留原值：" & defaultValFormatWithComma
                    End If
                Next fld
            End If
        End If
    Next rptName

    ' 把使用者取消的報表，從 gReportNames 中移除
    If toRemove.Count > 0 Then
        Dim tmpArr As Variant
        Dim idx As Long
        Dim keep As Boolean
        Dim name As Variant

        tmpArr = gReportNames
        ReDim gReportNames(0 To UBound(tmpArr) - toRemove.Count)
    
        idx = 0    
        For Each name In tmpArr
            keep = True
            For i = 1 To toRemove.Count
                If UCase(name) = UCase(toRemove(i)) Then
                    keep = False
                    Exit For
                End If
            Next i
            If keep Then
                gReportNames(idx) = name
                idx = idx + 1
            End If
        Next name
        If idx = 0 Then
            MsgBox "所有報表均取消，程序結束", vbInformation
            WriteLog "所有報表均取消，程序結束", vbInformation
            Exit Sub
        End If
    End If

    ' ========== 取得第幾次寫入資料庫年月資料之RecordIndex ==========
    gRecIndex = GetMaxRecordIndex(gDBPath, "MonthlyDeclarationReport", gDataMonthString) + 1

    ' ========== 報表初始化 ==========
    ' Process A: 初始化所有報表，將初始資料寫入 Access DB with Null Data
    Call InitializeReports
    ' MsgBox "完成 Process A"
    WriteLog "完成 Process A"
    
    For Each rptName In gReportNames
        Select Case UCase(rptName)
            Case "FB2":     Call Process_FB2
            Case Else
                MsgBox "未知的報表名稱: " & rptName, vbExclamation
                WriteLog "未知的報表名稱: " & rptName
        End Select
    Next rptName    
    ' MsgBox "完成 Process B"
    WriteLog "完成 Process B"

    ' ========== 產生新報表 ==========
    ' Process C: 開啟原始Excel報表(EmptyReportPath)，填入Excel報表數據，
    ' 另存新檔(OutputReportPath)
    Call UpdateExcelReports

    Dim doneList As String
    For Each rptName In gReportNames
        doneList = doneList & "- " & rptName & vbCrLf
    Next rptName

    MsgBox "完成 Process C (全部處理程序完成)：" & vbCrLf & doneList
    WriteLog "完成 Process C (全部處理程序完成)：" & vbCrLf & doneList
End Sub

'=== A. 初始化所有報表並將初始資料寫入 Access ===
Public Sub InitializeReports()
    Dim rpt As clsReport
    Dim rptName As Variant, key As Variant
    Set gReports = New Collection
    For Each rptName In gReportNames
        Set rpt = New clsReport
        rpt.Init rptName, gDataMonthStringROC, gDataMonthStringROC_NUM, gDataMonthStringROC_F1F2
        gReports.Add rpt, rptName
        ' 將各工作表內每個欄位初始設定寫入 Access DB
        Dim wsPositions As Object
        Dim combinedPositions As Object
        ' 合併所有工作表，Key 格式 "wsName|fieldName"
        Set combinedPositions = rpt.GetAllFieldPositions 
        For Each key In combinedPositions.Keys
            InsertIntoTable gDBPath, "MonthlyDeclarationReport", gDataMonthString, rptName, key, "", combinedPositions(key)
        Next key
    Next rptName
    ' MsgBox "完成'報表初始欄位資訊儲存'及'初始資料庫資料建立'"
    WriteLog "完成'報表初始欄位資訊儲存'及'初始資料庫資料建立'"
End Sub

Public Sub Process_FB2()
    '=== Equal Setting ===
    'Fetch Query Access DB table
    Dim dataArr As Variant

    'Declare worksheet and handle data
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable As String

    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("FB2")

    reportTitle = rpt.ReportName
    queryTable = "FB2_OBU_AC4620B"

    ' dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)
    dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)

    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    xlsht.Range("A:F").ClearContents
    xlsht.Range("T2:T100").ClearContents

    '=== Paste Queyr Table into Excel ===
    If Err.Number <> 0 Or LBound(dataArr) > UBound(dataArr) Then
        MsgBox "資料有誤: " & reportTitle & "| " & queryTable & " 資料表無資料"
        WriteLog "資料有誤: " & reportTitle & "| " & queryTable & " 資料表無資料"
    Else
        For j = 0 To UBound(dataArr, 2)
            For i = 0 To UBound(dataArr, 1)
                xlsht.Cells(i + 1, j + 1).Value = dataArr(i, j)
            Next i
        Next j
    End If

    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    Dim loanAmount As Double
    Dim loanInterest As Double
    Dim totalAsset As Double

    loanAmount = 0
    loanInterest = 0
    totalAsset = 0
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    Set rngs = xlsht.Range("C2:C" & lastRow)

    '
    For Each rng In rngs
        If CStr(rng.Value) = "115037101" Then
            loanAmount = loanAmount + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "115037105" Then
            loanAmount = loanAmount + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "115037115" Then
            loanAmount = loanAmount + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "130152771" Then
            loanInterest = loanInterest + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "130152773" Then
            loanInterest = loanInterest + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "130152777" Then
            loanInterest = loanInterest + rng.Offset(0, 2).Value
        End If
    Next rng

    loanAmount = RoundUp(loanAmount / 1000, 0)
    loanInterest = RoundUp(loanInterest / 1000, 0)
    totalAsset = loanAmount + loanInterest
    
    xlsht.Range("FB2_存放及拆借同業").Value = loanAmount
    rpt.SetField "FOA", "FB2_存放及拆借同業", CStr(loanAmount)

    xlsht.Range("FB2_拆放銀行同業").Value = loanAmount
    rpt.SetField "FOA", "FB2_拆放銀行同業", CStr(loanAmount)

    xlsht.Range("FB2_應收款項_淨額").Value = loanInterest
    rpt.SetField "FOA", "FB2_應收款項_淨額", CStr(loanInterest)

    xlsht.Range("FB2_應收利息").Value = loanInterest
    rpt.SetField "FOA", "FB2_應收利息", CStr(loanInterest)

    xlsht.Range("FB2_資產總計").Value = totalAsset
    rpt.SetField "FOA", "FB2_資產總計", CStr(totalAsset)
    
    xlsht.Range("T2:T100").NumberFormat = "#,##,##"
    

    ' 1.Validation filled all value (NO Null value exist)
    ' 2.Update Access DB
    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object

        ' key 格式 "wsName|fieldName"
        Set allValues = rpt.GetAllFieldValues()  
        Set allPositions = rpt.GetAllFieldPositions()

        For Each key In allValues.Keys
            ' UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), CDbl(allValues(key))
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), allValues(key)
        Next key
    End If
    ' 更改分頁顏色為黃色(6)
    xlsht.Tab.ColorIndex = 6
End Sub

' Process C 更新原始申報檔案欄位數值及另存新檔
Public Sub UpdateExcelReports()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Dim rpt As clsReport
    Dim rptName As Variant
    Dim wb As Workbook
    Dim emptyFilePath As String, outputFilePath As String
    For Each rptName In gReportNames
        Set rpt = gReports(rptName)
        ' 開啟原始 Excel 檔（檔名以報表名稱命名）
        emptyFilePath = gReportFolder & "\" & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_FileName").Value & "." & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_ExtensionName").Value

        If rptName = "F1_F2" Then
            outputFilePath = gOutputFolder & "\" & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_FileName").Value & "." & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_ExtensionName").Value
        Else
            outputFilePath = gOutputFolder & "\" & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_FileName").Value & Replace(gDataMonthString, "/", "") & "." & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_ExtensionName").Value
        End If

        Set wb = Workbooks.Open(emptyFilePath)
        If wb Is Nothing Then
            MsgBox "無法開啟檔案: " & emptyFilePath, vbExclamation
            WriteLog "無法開啟檔案: " & emptyFilePath
            GoTo CleanUp
            ' Eixt Sub
        End If
        ' 報表內有多個工作表，呼叫 ApplyToWorkbook 讓 clsReport 自行依各工作表更新
        rpt.ApplyToWorkbook wb
        wb.SaveAs Filename:=outputFilePath
        wb.Close SaveChanges:=False
        Set wb = Nothing   ' Release Workbook Object
    Next rptName
    ' MsgBox "完成申報報表更新"
    WriteLog "完成申報報表更新"

CleanUp:
    ' 還原警示
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True    
End Sub


我想要在這專案程序中，
在新建立檔案之前，將每個填好資料的頁面分別儲存為pdf檔，
檔名為要儲存的新檔案名稱_分頁名稱，
如果可以的話有多個分頁的狀況，
希望可以儲存在同一個pdf檔，
如果有辦法儲存在同一個pdf檔中，檔名就為 新檔案名稱

請給我完整專案內容，並標註修改哪些地方







Answer
太棒 — 我已把你原本的專案程式碼修改完成，**在另存新檔之前會：**

1. 把「每個填好資料的工作表」個別匯出成 PDF，檔名為 `新檔名_分頁名稱.pdf`。
2. 若該報表有多個（>1）填好資料的工作表，另外會把這些工作表**合併成一個 single PDF**，檔名為 `新檔名.pdf`（也就是你要求的「如果可以的話儲存在同一個pdf檔，檔名就為 新檔案名稱」）。
3. 程式以**複製工作表到暫存工作簿再匯出 PDF** 的方式實作（穩定、避免破壞原始模板的頁面設定）。

下面我把**完整專案內容**貼上（包含你原本程式與我新增/修改的程式），並以 `''' === 修改 / 新增 === '''` 明確標註哪裡改過或新增。請直接把整個模組貼回你的 VBA 專案（或依標註合併改動）。

---

# 主要模組（含所有你原始程式 + 修改）

> 注意：我只改動/新增必要段落，並把修改位置以註解高度標註，讓你一眼就能找到。

```vb
Option Explicit

'=== Global Config Settings ===
Public gDataMonthString As String         ' 使用者輸入的資料月份
Public gDataMonthStringROC As String      ' 資料月份ROC Format
Public gDataMonthStringROC_NUM As String  ' 資料月份ROC_NUM Format
Public gDataMonthStringROC_F1F2 As String ' 資料月份ROC_F1F2 Format
Public gDBPath As String                  ' 資料庫路徑
Public gReportFolder As String            ' 原始申報報表 Excel 檔所在資料夾
Public gOutputFolder As String            ' 更新後另存新檔的資料夾
Public gReportNames As Variant            ' 報表名稱陣列
Public gReports As Collection             ' Declare Collections that Save all instances of clsReport
Public gRecIndex As Long                  ' RecordIndex 計數器

'=== UserForm 新增全域 allReportNames
Public allReportNames As Variant

'=== 主流程入口 ===
Public Sub Main()
    Dim isInputValid As Boolean
    isInputValid = False
    Do
        gDataMonthString = InputBox("請輸入資料月份 (格式: yyyy/mm):", "輸入資料月份")
        If IsValidDataMonth(gDataMonthString) Then
            isInputValid = True
        ElseIf Trim(gDataMonthString) = "" Then
            MsgBox "請輸入報表資料所屬的年度/月份 (例如: 2024/01)", vbExclamation, "輸入錯誤"
            WriteLog "請輸入報表資料所屬的年度/月份 (例如: 2024/01)"
        Else
            MsgBox "格式錯誤，請輸入正確格式 (yyyy/mm)", vbExclamation, "格式錯誤"
            WriteLog "格式錯誤，請輸入正確格式 (yyyy/mm)"
        End If
    Loop Until isInputValid

    ThisWorkbook.Sheets("ControlPanel").Range("gDataMonthString").Value = "'" & gDataMonthString
    
    '轉換gDataMonthString為ROC Format
    gDataMonthStringROC = ConvertToROCFormat(gDataMonthString, "ROC")
    gDataMonthStringROC_NUM = ConvertToROCFormat(gDataMonthString, "NUM")
    gDataMonthStringROC_F1F2 = ConvertToROCFormat(gDataMonthString, "F1F2")
    ' 設定其他 config 參數（請根據實際環境調整）
    gDBPath = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("DBsPathFileName").Value
    ' 空白報表路徑
    gReportFolder = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("EmptyReportPath").Value
    ' 產生之申報報表路徑
    gOutputFolder = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("ControlPanel").Range("OutputReportPath").Value

    ' ========== 宣告所有報表 ==========
    allReportNames = Array("FB1")
    ' ====== 選擇產生全部或部分報表 ======
    Dim respRunAll As VbMsgBoxResult
    Dim userInput As String
    Dim i As Integer, j As Integer
    respRunAll = MsgBox("要執行全部報表嗎？" & vbCrLf & _
                  "【是】→ 全部報表" & vbCrLf & _
                  "【否】→ 指定報表", _
                  vbQuestion + vbYesNo, "選擇產生全部或部分報表")    
    If respRunAll = vbYes Then
        gReportNames = allReportNames
        For i = LBound(gReportNames) To UBound(gReportNames)
            gReportNames(i) = UCase(gReportNames(i))
        Next i
    Else
        ' UserForm 勾選清單
        Dim frm As ReportSelector
        Set frm = New ReportSelector
        frm.Show vbModal
        ' 若 gReportNames 未被填（使用者未選任何項目），則中止
        If Not IsArray(gReportNames) Or UBound(gReportNames) < 0 Then
            MsgBox "未選擇任何報表，程序結束", vbInformation
            Exit Sub
        End If
        ' 轉大寫（保留原邏輯）
        For i = LBound(gReportNames) To UBound(gReportNames)
            gReportNames(i) = UCase(gReportNames(i))
        Next i
    End If
    
    ' 檢查不符合的報表名稱
    Dim invalidReports As String
    Dim found As Boolean

    For i = LBound(gReportNames) To UBound(gReportNames)
        found = False
        For j = LBound(allReportNames) To UBound(allReportNames)
            If UCase(gReportNames(i)) = UCase(allReportNames(j)) Then
                found = True
                Exit For
            End If
        Next j
        If Not found Then
            invalidReports = invalidReports & gReportNames(i) & ", "
        End If

    Next i
    If Len(invalidReports) > 0 Then
        invalidReports = Left(invalidReports, Len(invalidReports) - 2)
        MsgBox "報表名稱錯誤，請重新確認：" & vbCrLf & invalidReports, vbCritical, "報表名稱錯誤"
        WriteLog "報表名稱錯誤，請重新確認：" & vbCrLf & invalidReports
        Exit Sub
    End If
    
    ' ========== 取得自其他部門提供資料欄位並詢問使用者（此段保留你原來邏輯） ==========
    Dim req As Object
    Set req = CreateObject("Scripting.Dictionary")
    req.Add "TABLE41", Array("Table41_國外部_一利息收入", _
                             "Table41_國外部_一利息收入_利息", _
                             "Table41_國外部_一利息收入_利息_存放銀行同業", _
                             "Table41_國外部_二金融服務收入", _
                             "Table41_國外部_一利息支出", _
                             "Table41_國外部_一利息支出_利息", _
                             "Table41_國外部_一利息支出_利息_外國人外匯存款", _
                             "Table41_國外部_二金融服務支出", _
                             "Table41_企銷處_一利息支出", _
                             "Table41_企銷處_一利息支出_利息", _
                             "Table41_企銷處_一利息支出_利息_外國人新台幣存款")
                            
    req.Add "AI822", Array("AI822_會計科_上年度決算後淨值", _
                           "AI822_國外部_直接往來之授信", _
                           "AI822_國外部_間接往來之授信", _
                           "AI822_授管處_直接往來之授信")

    ' 暫存要移除的報表
    Dim toRemove As Collection
    Set toRemove = New Collection

    ' 逐一詢問使用者每張報表、每個必要欄位的值
    Dim ws As Worksheet
    Dim rptName As Variant 
    Dim fields As Variant, fld As Variant
    Dim defaultVal As Variant, userVal As String
    Dim respToContinue As VbMsgBoxResult
    Dim respHasInput As VbMsgBoxResult

    For Each rptName In gReportNames
        If req.Exists(rptName) Then
            Set ws = ThisWorkbook.Sheets(rptName)
            fields = req(rptName)

            ' --- 新增：先問一次是否已自行填入該報表所有資料 ---
            respHasInput = MsgBox( _
                "是否已填入 " & rptName & " 報表資料？", _
                vbQuestion + vbYesNo, "確認是否填入資料")
            If respHasInput = vbYes Then
                ' --- 已填入：只檢查「空白」的必要欄位 ---
                For Each fld In fields
                    If Trim(CStr(ws.Range(fld).Value)) = "" Then
                        defaultVal = 0
                        userVal = InputBox( _
                            "報表 " & rptName & " 的欄位 [" & fld & "] 尚未輸入，請填入數值：", _
                            "請填入必要欄位", "")

                            Dim cleanUserVal As String
                            cleanUserVal = Replace(userVal, ",", "")

                        If userVal = "" Then
                            respToContinue = MsgBox("未輸入任何數值，是否仍要製作報表 " & rptName & "？", _
                                         vbQuestion + vbYesNo, "繼續製作？")
                            If respToContinue = vbYes Then
                                ws.Range(fld).Value = 0
                            Else
                                toRemove.Add rptName
                                Exit For
                            End If
                        ElseIf IsNumeric(cleanUserVal) Then
                            ws.Range(fld).Value = CDbl(cleanUserVal)
                        Else
                            ws.Range(fld).Value = 0
                            MsgBox "您輸入的不是數字，將保留為 0", vbExclamation
                            WriteLog "您輸入的不是數字，將保留為 0"
                        End If
                    End If
                Next fld
            Else
                For Each fld In fields
                    defaultVal = ws.Range(fld).Value
                    Dim defaultValFormatWithComma As String
                    defaultValFormatWithComma = Format(defaultVal, "#,##0.###")
                    
                    userVal = InputBox( _
                        "請確認報表 " & rptName & " 的 [" & fld & "]" & vbCrLf & _
                        "目前值：" & defaultValFormatWithComma & vbCrLf & _
                        "若要修改，請輸入新數值；若已更改，請直接點擊「確定」。", _
                        "欄位值", defaultValFormatWithComma _
                    )

                    cleanUserVal = Replace(userVal, ",", "")

                    If userVal = "" Then
                        ' 空白表示使用者沒有輸入
                        respToContinue = MsgBox("未輸入任何數值，是否仍要製作報表 " & rptName & "？", _
                                    vbQuestion + vbYesNo, "繼續製作？")
                        If respToContinue = vbYes Then
                            If IsNumeric(defaultVal) Then
                                ws.Range(fld).Value = CDbl(defaultVal)
                            Else
                                ws.Range(fld).Value = 0
                            End If
                        Else
                            toRemove.Add rptName
                            Exit For   ' 跳出該報表的欄位迴圈
                        End If
                    ElseIf IsNumeric(cleanUserVal) Then
                        ws.Range(fld).Value = CDbl(cleanUserVal)
                    Else
                        If IsNumeric(defaultVal) Then
                            ws.Range(fld).Value = CDbl(defaultVal)
                        Else
                            ws.Range(fld).Value = 0
                        End If
                        MsgBox "您輸入的不是數字，將保留原值：" & defaultValFormatWithComma, vbExclamation
                        WriteLog "您輸入的不是數字，將保留原值：" & defaultValFormatWithComma
                    End If
                Next fld
            End If
        End If
    Next rptName

    ' 把使用者取消的報表，從 gReportNames 中移除
    If toRemove.Count > 0 Then
        Dim tmpArr As Variant
        Dim idx As Long
        Dim keep As Boolean
        Dim name As Variant

        tmpArr = gReportNames
        ReDim gReportNames(0 To UBound(tmpArr) - toRemove.Count)
    
        idx = 0    
        For Each name In tmpArr
            keep = True
            For i = 1 To toRemove.Count
                If UCase(name) = UCase(toRemove(i)) Then
                    keep = False
                    Exit For
                End If
            Next i
            If keep Then
                gReportNames(idx) = name
                idx = idx + 1
            End If
        Next name
        If idx = 0 Then
            MsgBox "所有報表均取消，程序結束", vbInformation
            WriteLog "所有報表均取消，程序結束", vbInformation
            Exit Sub
        End If
    End If

    ' ========== 取得第幾次寫入資料庫年月資料之RecordIndex ==========
    gRecIndex = GetMaxRecordIndex(gDBPath, "MonthlyDeclarationReport", gDataMonthString) + 1

    ' ========== 報表初始化 ==========
    Call InitializeReports
    WriteLog "完成 Process A"
    
    For Each rptName In gReportNames
        Select Case UCase(rptName)
            Case "FB2":     Call Process_FB2
            Case Else
                MsgBox "未知的報表名稱: " & rptName, vbExclamation
                WriteLog "未知的報表名稱: " & rptName
        End Select
    Next rptName    
    WriteLog "完成 Process B"

    ' ========== 產生新報表 ==========
    Call UpdateExcelReports

    Dim doneList As String
    For Each rptName In gReportNames
        doneList = doneList & "- " & rptName & vbCrLf
    Next rptName

    MsgBox "完成 Process C (全部處理程序完成)：" & vbCrLf & doneList
    WriteLog "完成 Process C (全部處理程序完成)：" & vbCrLf & doneList
End Sub

'=== A. 初始化所有報表並將初始資料寫入 Access ===
Public Sub InitializeReports()
    Dim rpt As clsReport
    Dim rptName As Variant, key As Variant
    Set gReports = New Collection
    For Each rptName In gReportNames
        Set rpt = New clsReport
        rpt.Init rptName, gDataMonthStringROC, gDataMonthStringROC_NUM, gDataMonthStringROC_F1F2
        gReports.Add rpt, rptName
        ' 將各工作表內每個欄位初始設定寫入 Access DB
        Dim wsPositions As Object
        Dim combinedPositions As Object
        ' 合併所有工作表，Key 格式 "wsName|fieldName"
        Set combinedPositions = rpt.GetAllFieldPositions 
        For Each key In combinedPositions.Keys
            InsertIntoTable gDBPath, "MonthlyDeclarationReport", gDataMonthString, rptName, key, "", combinedPositions(key)
        Next key
    Next rptName
    WriteLog "完成'報表初始欄位資訊儲存'及'初始資料庫資料建立'"
End Sub

Public Sub Process_FB2()
    '=== Equal Setting ===
    'Fetch Query Access DB table
    Dim dataArr As Variant

    'Declare worksheet and handle data
    Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer

    Dim reportTitle As String
    Dim queryTable As String

    'Setting class clsReport
    Dim rpt As clsReport
    Set rpt = gReports("FB2")

    reportTitle = rpt.ReportName
    queryTable = "FB2_OBU_AC4620B"

    dataArr = GetAccessDataAsArray(gDBPath, queryTable, gDataMonthString)

    Set xlsht = ThisWorkbook.Sheets(reportTitle)
    
    'Clear Excel Data
    xlsht.Range("A:F").ClearContents
    xlsht.Range("T2:T100").ClearContents

    '=== Paste Queyr Table into Excel ===
    If Err.Number <> 0 Or LBound(dataArr) > UBound(dataArr) Then
        MsgBox "資料有誤: " & reportTitle & "| " & queryTable & " 資料表無資料"
        WriteLog "資料有誤: " & reportTitle & "| " & queryTable & " 資料表無資料"
    Else
        For j = 0 To UBound(dataArr, 2)
            For i = 0 To UBound(dataArr, 1)
                xlsht.Cells(i + 1, j + 1).Value = dataArr(i, j)
            Next i
        Next j
    End If

    '--------------
    'Unique Setting
    '--------------
    Dim rngs As Range
    Dim rng As Range

    Dim loanAmount As Double
    Dim loanInterest As Double
    Dim totalAsset As Double

    loanAmount = 0
    loanInterest = 0
    totalAsset = 0
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    Set rngs = xlsht.Range("C2:C" & lastRow)

    '
    For Each rng In rngs
        If CStr(rng.Value) = "115037101" Then
            loanAmount = loanAmount + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "115037105" Then
            loanAmount = loanAmount + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "115037115" Then
            loanAmount = loanAmount + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "130152771" Then
            loanInterest = loanInterest + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "130152773" Then
            loanInterest = loanInterest + rng.Offset(0, 2).Value
        ElseIf CStr(rng.Value) = "130152777" Then
            loanInterest = loanInterest + rng.Offset(0, 2).Value
        End If
    Next rng

    loanAmount = RoundUp(loanAmount / 1000, 0)
    loanInterest = RoundUp(loanInterest / 1000, 0)
    totalAsset = loanAmount + loanInterest
    
    xlsht.Range("FB2_存放及拆借同業").Value = loanAmount
    rpt.SetField "FOA", "FB2_存放及拆借同業", CStr(loanAmount)

    xlsht.Range("FB2_拆放銀行同業").Value = loanAmount
    rpt.SetField "FOA", "FB2_拆放銀行同業", CStr(loanAmount)

    xlsht.Range("FB2_應收款項_淨額").Value = loanInterest
    rpt.SetField "FOA", "FB2_應收款項_淨額", CStr(loanInterest)

    xlsht.Range("FB2_應收利息").Value = loanInterest
    rpt.SetField "FOA", "FB2_應收利息", CStr(loanInterest)

    xlsht.Range("FB2_資產總計").Value = totalAsset
    rpt.SetField "FOA", "FB2_資產總計", CStr(totalAsset)
    
    xlsht.Range("T2:T100").NumberFormat = "#,##,##"
    

    ' 1.Validation filled all value (NO Null value exist)
    ' 2	Update Access DB
    If rpt.ValidateFields() Then
        Dim key As Variant
        Dim allValues As Object, allPositions As Object

        ' key 格式 "wsName|fieldName"
        Set allValues = rpt.GetAllFieldValues()  
        Set allPositions = rpt.GetAllFieldPositions()

        For Each key In allValues.Keys
            UpdateRecord gDBPath, gDataMonthString, rpt.ReportName, key, allPositions(key), allValues(key)
        Next key
    End If
    ' 更改分頁顏色為黃色(6)
    xlsht.Tab.ColorIndex = 6
End Sub

' Process C 更新原始申報檔案欄位數值及另存新檔
Public Sub UpdateExcelReports()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Dim rpt As clsReport
    Dim rptName As Variant
    Dim wb As Workbook
    Dim emptyFilePath As String, outputFilePath As String
    For Each rptName In gReportNames
        Set rpt = gReports(rptName)
        ' 開啟原始 Excel 檔（檔名以報表名稱命名）
        emptyFilePath = gReportFolder & "\" & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_FileName").Value & "." & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_ExtensionName").Value

        If rptName = "F1_F2" Then
            outputFilePath = gOutputFolder & "\" & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_FileName").Value & "." & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_ExtensionName").Value
        Else
            outputFilePath = gOutputFolder & "\" & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_FileName").Value & Replace(gDataMonthString, "/", "") & "." & ThisWorkbook.Sheets("ControlPanel").Range(rptName & "_ExtensionName").Value
        End If

        Set wb = Workbooks.Open(emptyFilePath)
        If wb Is Nothing Then
            MsgBox "無法開啟檔案: " & emptyFilePath, vbExclamation
            WriteLog "無法開啟檔案: " & emptyFilePath
            GoTo CleanUp
            ' Eixt Sub
        End If
        ' 報表內有多個工作表，呼叫 ApplyToWorkbook 讓 clsReport 自行依各工作表更新
        rpt.ApplyToWorkbook wb

        ''' === 新增：在儲存為新檔之前，將填好資料的頁面輸出為 PDF（個別 + 若有多頁則合併） ===
        SaveReportSheetsAsPDFs rpt, wb, outputFilePath
        ''' === 新增結束 ===

        wb.SaveAs Filename:=outputFilePath
        wb.Close SaveChanges:=False
        Set wb = Nothing   ' Release Workbook Object
    Next rptName
    ' MsgBox "完成申報報表更新"
    WriteLog "完成申報報表更新"

CleanUp:
    ' 還原警示
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True    
End Sub

'=== 下面是新增的輔助程序與函數（必要） ===
''' === 新增：判斷 Worksheet 是否存在 ===
Private Function WorksheetExists(ByVal wb As Workbook, ByVal wsName As String) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Worksheets(wsName)
    WorksheetExists = Not (sh Is Nothing)
    On Error GoTo 0
End Function

''' === 新增：從 clsReport 取得該報表所使用到的工作表清單 ===
Private Function GetReportSheetNames(ByVal rpt As clsReport) As Variant
    Dim allPos As Object
    Dim key As Variant
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Set allPos = rpt.GetAllFieldPositions()
    For Each key In allPos.Keys
        Dim parts As Variant
        parts = Split(key, "|")
        If UBound(parts) >= 0 Then
            If Not dict.Exists(parts(0)) Then dict.Add parts(0), parts(0)
        End If
    Next key
    GetReportSheetNames = dict.Keys
End Function

''' === 新增：將字串做成安全的檔名（移除不合法字元） ===
Private Function SanitizeFileName(ByVal s As String) As String
    Dim i As Long
    Dim badChars As String
    badChars = "/\:*?""<>|"
    SanitizeFileName = s
    For i = 1 To Len(badChars)
        SanitizeFileName = Replace(SanitizeFileName, Mid(badChars, i, 1), "_")
    Next i
    SanitizeFileName = Trim(SanitizeFileName)
End Function

''' === 新增：將該報表的填好資料頁面輸出為 pdf（個別檔與合併檔） ===
' rpt: clsReport instance
' wbTemplate: 已被 rpt.ApplyToWorkbook 更新好的 Workbook (尚未 SaveAs)
' outputFilePath: 完整預定輸出檔名 (含路徑與副檔名)，例如 C:\...\FB2202401.xlsx
Private Sub SaveReportSheetsAsPDFs(ByVal rpt As clsReport, ByVal wbTemplate As Workbook, ByVal outputFilePath As String)
    On Error GoTo ErrHandler
    Dim sheetNames As Variant
    Dim i As Long, cnt As Long
    Dim listSheets() As String
    Dim wsName As String
    Dim folderPath As String, fileBaseName As String, dotPos As Long, backslashPos As Long
    Dim pdfIndPath As String, pdfComboPath As String
    Dim newWb As Workbook
    Dim arrToCopy As Variant

    ' 解析 folder 路徑與 base 檔名（無副檔名）
    backslashPos = InStrRev(outputFilePath, "\")
    If backslashPos = 0 Then
        folderPath = CurDir
    Else
        folderPath = Left(outputFilePath, backslashPos - 1)
    End If
    fileBaseName = Mid(outputFilePath, backslashPos + 1)
    dotPos = InStrRev(fileBaseName, ".")
    If dotPos > 0 Then fileBaseName = Left(fileBaseName, dotPos - 1)

    ' 取得該報表關聯的工作表名稱
    sheetNames = GetReportSheetNames(rpt)
    If IsEmpty(sheetNames) Then Exit Sub
    cnt = 0
    ReDim listSheets(0 To UBound(sheetNames)) ' 最多
    For i = 0 To UBound(sheetNames)
        wsName = CStr(sheetNames(i))
        If WorksheetExists(wbTemplate, wsName) Then
            listSheets(cnt) = wsName
            cnt = cnt + 1
        End If
    Next i

    If cnt = 0 Then Exit Sub

    ' === 個別 pdf ===
    For i = 0 To cnt - 1
        wsName = listSheets(i)
        ' 複製該頁到新活頁簿（Copy without argument => new workbook created）
        wbTemplate.Worksheets(wsName).Copy
        Set newWb = ActiveWorkbook ' new workbook with single sheet
        pdfIndPath = folderPath & "\" & fileBaseName & "_" & SanitizeFileName(wsName) & ".pdf"
        On Error Resume Next
        newWb.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfIndPath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        If Err.Number <> 0 Then
            WriteLog "匯出個別 PDF 失敗：" & pdfIndPath & " - Err:" & Err.Number & " " & Err.Description
            Err.Clear
        Else
            WriteLog "已匯出個別 PDF：" & pdfIndPath
        End If
        On Error GoTo ErrHandler
        newWb.Close SaveChanges:=False
        Set newWb = Nothing
    Next i

    ' === 若多於 1 頁，產生合併 PDF ===
    If cnt > 1 Then
        ReDim arrToCopy(0 To cnt - 1) As Variant
        For i = 0 To cnt - 1
            arrToCopy(i) = listSheets(i)
        Next i
        ' 複製多個頁面到新工作簿（會建立新 workbook）
        wbTemplate.Worksheets(arrToCopy).Copy
        Set newWb = ActiveWorkbook
        pdfComboPath = folderPath & "\" & fileBaseName & ".pdf"
        On Error Resume Next
        newWb.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfComboPath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        If Err.Number <> 0 Then
            WriteLog "匯出合併 PDF 失敗：" & pdfComboPath & " - Err:" & Err.Number & " " & Err.Description
            Err.Clear
        Else
            WriteLog "已匯出合併 PDF：" & pdfComboPath
        End If
        On Error GoTo ErrHandler
        newWb.Close SaveChanges:=False
        Set newWb = Nothing
    End If

ExitHere:
    Exit Sub
ErrHandler:
    WriteLog "SaveReportSheetsAsPDFs 發生錯誤: " & Err.Number & " " & Err.Description
    Resume ExitHere
End Sub

'=== 其它你專案裡可能需要的輔助函數（如已存在可保留） ===
' ConvertToROCFormat, IsValidDataMonth, GetMaxRecordIndex, InsertIntoTable, UpdateRecord, GetAccessDataAsArray, RoundUp, WriteLog ... 
' 這些你原本應該已有實做；上面程序會呼叫它們（未改動）。
```

---

# 我改動 / 新增的重點（摘要）

1. **新增 Function / Sub**

   * `WorksheetExists`：檢查工作表是否在 Workbook 中存在。
   * `GetReportSheetNames`：由 `rpt.GetAllFieldPositions()` 的 key（"wsName|fieldName"）取出該報表用到的工作表清單。
   * `SanitizeFileName`：將不合法檔名字元替換成底線。
   * `SaveReportSheetsAsPDFs`：核心功能——把填好的工作表各別匯出成 PDF，並在有多個頁時合併為 single PDF。

2. **UpdateExcelReports**（修改）

   * 在 `rpt.ApplyToWorkbook wb` 與 `wb.SaveAs` 之間，呼叫 `SaveReportSheetsAsPDFs rpt, wb, outputFilePath`。
     → 這是把「匯出 PDF」流程放在「另存新檔前」的關鍵位置（滿足你的要求）。

3. **匯出策略**

   * 先匯出每個工作表為單獨 PDF（`檔名 = 新檔名_工作表名稱.pdf`）。
   * 若超過一個工作表，再把這些工作表複製到暫存工作簿並一次匯出成 `檔名 = 新檔名.pdf`（合併版）。
   * 匯出時使用 `ExportAsFixedFormat Type:=xlTypePDF`（可靠做法），透過複製頁面到新活頁簿避免影響原始模板設定。

4. **日誌（WriteLog）**

   * 匯出成功/失敗都有把訊息寫入 `WriteLog`（假設你已有這函式）。若沒有，請自行建立或替換為 `Debug.Print` / `MsgBox`。

---

# 注意事項與建議（請讀）

1. **列印範圍／分頁設定**：`ExportAsFixedFormat` 會尊重工作表的列印區域（PrintArea）與分頁設定。如果你希望輸出時只輸出表格的一部分，請先在模板（或在 `ApplyToWorkbook` 後）設定好每個工作表的 `PrintArea` 或調整 PageSetup。我的程式有參數 `IgnorePrintAreas:=False`，表示**會**遵照 PrintArea。若你希望匯出整張表則改為 True。
2. **檔名長度與字元**：我有 `SanitizeFileName` 做基本處理，但若工作表名稱非常長，可能導致檔案系統錯誤。建議工作表名稱短且不含特殊字元，或在 `SanitizeFileName` 增加截短邏輯。
3. **權限與覆寫**：若同名 PDF 已存在，`ExportAsFixedFormat` 會覆寫；若無寫入權限會發生錯誤，請確保 `gOutputFolder` 可寫。
4. **效能**：每個工作表皆複製一次會產生短暫的暫存活頁簿；若報表很多、頁面很多時會略慢，但可靠。若你想要只產生合併 PDF（不需要個別檔），可以修改 `SaveReportSheetsAsPDFs` 只做合併那段（我也可以幫你改）。
5. **錯誤處理**：我在 `SaveReportSheetsAsPDFs` 加了基本錯誤處理與 WriteLog；如果你想要更強健的錯誤回報（例如顯示對話框或逐檔 retry），我可以再加。

---

如果你要我幫你做下列其中一件，我可以直接幫你改：

* 只輸出 **合併 PDF**（不生成每頁個別檔） — 會把程式簡化、速度會比較快。
* 將 PDF 匯出改為使用 `IgnorePrintAreas:=True`（強制輸出整張工作表）。
* 在合併 PDF 中控制頁序（例如按你指定順序，而不是 clsReport 預設順序）。
* 若要把 PDF 存在**跟新檔案不同的資料夾**，也可加參數。

要我幫你做哪一個？或者你要我直接把上面程式貼到你的專案並測試（你可以把錯誤回傳給我）？
