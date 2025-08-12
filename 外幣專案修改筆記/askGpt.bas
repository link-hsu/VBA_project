
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

' ========


1.製作完申報Excel表時，產生一個pdf檔案
2.提供一個參數為儲存pdf的Root Path

' ========

Control Panel

1.將搜尋篩選過後的pdf貼入PDFCreator資料夾中
2.將資料貼入人工製作報表中，call相關ButtonSub執行程序
3.產生輔助報表pdf檔案

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\AI
D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\AI\AI230
D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\AI\AI230\申報報表
D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\AI\AI230\申報報表\202506

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\AI\AI240

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\AI\AI821

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\AI\AI822申報檔

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\CNY

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\F1F2

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\F1F2\用表

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\F1F2\申報檔

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\FB

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\FM

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\FM\FM2

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\FM\FM5





###寫一個函數 input 為A檔案路徑、B檔案路徑、比對分頁、比對儲存格 (輸入類似這樣的 "C11:J36,C37:E47,G37:J47,F40,F42:F47,C48:C50,J48:J50,C52:C56,J52:J56,C57:J57")
我要將傳入的兩個檔案中同樣分頁同樣儲存格中的數值去逐一比對數值，
如果數值完全一致，回傳0或NULL
如果有數值不一致，回傳所有不一致的 (檔案名稱，分頁名稱，儲存格位置)

兩個檔案分頁和儲存格位置完全一致，但是其中的數值不同，我要寫出一個這樣的Function或Sub





' =====================

AI240

Option Explicit

Sub CopyDataToAI240()
    Dim wsDL6850 As Worksheet
    Dim wsAI240 As Worksheet
    Dim inputDate As Date
    Dim baseDate As Date
    Dim rowCount As Long
    Dim copyCount0To10 As Long
    Dim copyCount11To30 As Long
    Dim copyCount31To90 As Long
    Dim copyCount91To180 As Long
    Dim copyCount181To365 As Long
    Dim copyCount366To As Long
    Dim destRow0TO10 As Long
    Dim destRow11TO30 As Long
    Dim destRow31TO90 As Long
    Dim destRow91TO180 As Long
    Dim destRow181TO365 As Long
    Dim destRow366TO As Long
    Dim i As Long
    

    '設定工作表
    Set wsDL6850 = ThisWorkbook.Worksheets("DL6850原始資料")
    Set wsAI240 = ThisWorkbook.Worksheets("AI240")

    '讀取基準日期
    inputDate = InputBox("請輸入基準日(日期格式yyyy/mm/dd)：")
    baseDate = inputDate

    '填入基準日期至DL6850原始資料工作表和AI240工作表
    wsDL6850.Range("P1").Value = baseDate
    wsAI240.Range("A2").Value = baseDate
    
    
    '清空AI240工作表數據
    ' 清空範圍 A9:I58
    wsAI240.Range("A9:I58").ClearContents
    ' 清空範圍 L9:T58
    wsAI240.Range("L9:T58").ClearContents
    ' 清空範圍 A90:I139
    wsAI240.Range("A90:I139").ClearContents
    ' 清空範圍 L90:T139
    wsAI240.Range("L90:T139").ClearContents
    ' 清空範圍 A153:I162
    wsAI240.Range("A153:I162").ClearContents
    ' 清空範圍 L153:T162
    wsAI240.Range("L153:T162").ClearContents
    ' 清空範圍 A170:I179
    wsAI240.Range("A170:I179").ClearContents
    ' 清空範圍 L170:T179
    wsAI240.Range("L170:T179").ClearContents
    
    
    Call ImportDL6850CSV

    '刪除符合條件的資料（DL6850原始資料工作表 B欄位以及 E、H、C、J 欄位的條件）
    
    rowCount = wsDL6850.Cells(wsDL6850.Rows.Count, "B").End(xlUp).Row
    For i = rowCount To 2 Step -1
        If Left(wsDL6850.Range("B" & i).Value, 2) <> "TR" Then
            wsDL6850.Rows(i).Delete
        End If
    Next i
    
    rowCount = wsDL6850.Cells(wsDL6850.Rows.Count, "B").End(xlUp).Row
    For i = rowCount To 2 Step -1
        If (wsDL6850.Range("E" & i).Value <> "TWD" And wsDL6850.Range("H" & i).Value <> "TWD") _
        Or wsDL6850.Range("C" & i).Value <= baseDate _
        Or wsDL6850.Range("J" & i).Value > baseDate Then
            wsDL6850.Rows(i).Delete
        End If
    Next i
    
 

    '將符合條件的資料複製貼入AI240工作表
    rowCount = wsDL6850.Cells(wsDL6850.Rows.Count, "B").End(xlUp).Row
    
    
    
    ' SWOP(SS or SF) and OutFlow TWD(colH)
    '起始貼入的目標列
    destRow0TO10 = 9
    destRow11TO30 = 19
    destRow31TO90 = 29
    destRow91TO180 = 39
    destRow181TO365 = 49

    '初始化計數變數
    copyCount0To10 = 0
    copyCount11To30 = 0
    copyCount31To90 = 0
    copyCount91To180 = 0
    copyCount181To365 = 0

    For i = 2 To rowCount
        If (wsDL6850.Range("A" & i).Value Like "SS*" Or wsDL6850.Range("A" & i).Value Like "SF*") And wsDL6850.Range("H" & i).Value = "TWD" Then
            Select Case wsDL6850.Range("N" & i).Value
                Case 0 To 10
                    copyCount0To10 = copyCount0To10 + 1
                    If copyCount0To10 > 10 Then
                        MsgBox "此期間流出之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow0TO10 & ":I" & destRow0TO10).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow0TO10 = destRow0TO10 + 1

                Case 11 To 30
                    copyCount11To30 = copyCount11To30 + 1
                    If copyCount11To30 > 10 Then
                        MsgBox "此期間流出之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow11TO30 & ":I" & destRow11TO30).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow11TO30 = destRow11TO30 + 1

                Case 31 To 90
                    copyCount31To90 = copyCount31To90 + 1
                    If copyCount31To90 > 10 Then
                        MsgBox "此期間流出之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow31TO90 & ":I" & destRow31TO90).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow31TO90 = destRow31TO90 + 1

                Case 91 To 180
                    copyCount91To180 = copyCount91To180 + 1
                    If copyCount91To180 > 10 Then
                        MsgBox "此期間流出之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow91TO180 & ":I" & destRow91TO180).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow91TO180 = destRow91TO180 + 1

                Case 181 To 365
                    copyCount181To365 = copyCount181To365 + 1
                    If copyCount181To365 > 10 Then
                        MsgBox "此期間流出之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow181TO365 & ":I" & destRow181TO365).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow181TO365 = destRow181TO365 + 1
            End Select
        End If
    Next i
    
    
    
    ' SWOP(SS or SF) and InFlow TWD(colE)
    '起始貼入的目標列
    destRow0TO10 = 9
    destRow11TO30 = 19
    destRow31TO90 = 29
    destRow91TO180 = 39
    destRow181TO365 = 49

    '初始化計數變數
    copyCount0To10 = 0
    copyCount11To30 = 0
    copyCount31To90 = 0
    copyCount91To180 = 0
    copyCount181To365 = 0

    For i = 2 To rowCount
        If (wsDL6850.Range("A" & i).Value Like "SS*" Or wsDL6850.Range("A" & i).Value Like "SF*") And wsDL6850.Range("E" & i).Value = "TWD" Then
            Select Case wsDL6850.Range("N" & i).Value
                Case 0 To 10
                    copyCount0To10 = copyCount0To10 + 1
                    If copyCount0To10 > 10 Then
                        MsgBox "此期間流入之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow0TO10 & ":T" & destRow0TO10).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow0TO10 = destRow0TO10 + 1

                Case 11 To 30
                    copyCount11To30 = copyCount11To30 + 1
                    If copyCount11To30 > 10 Then
                        MsgBox "此期間流入之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow11TO30 & ":T" & destRow11TO30).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow11TO30 = destRow11TO30 + 1

                Case 31 To 90
                    copyCount31To90 = copyCount31To90 + 1
                    If copyCount31To90 > 10 Then
                        MsgBox "此期間流入之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow31TO90 & ":T" & destRow31TO90).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow31TO90 = destRow31TO90 + 1

                Case 91 To 180
                    copyCount91To180 = copyCount91To180 + 1
                    If copyCount91To180 > 10 Then
                        MsgBox "此期間流入之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow91TO180 & ":T" & destRow91TO180).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow91TO180 = destRow91TO180 + 1

                Case 181 To 365
                    copyCount181To365 = copyCount181To365 + 1
                    If copyCount181To365 > 10 Then
                        MsgBox "此期間流入之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow181TO365 & ":T" & destRow181TO365).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow181TO365 = destRow181TO365 + 1
            End Select
        End If
    Next i
    
    
    
    ' SPOT(FS) and OutFlow TWD(colH)
    '起始貼入的目標列
    destRow0TO10 = 90
    destRow11TO30 = 100
    destRow31TO90 = 110
    destRow91TO180 = 120
    destRow181TO365 = 130

    '初始化計數變數
    copyCount0To10 = 0
    copyCount11To30 = 0
    copyCount31To90 = 0
    copyCount91To180 = 0
    copyCount181To365 = 0

    For i = 2 To rowCount
        If wsDL6850.Range("A" & i).Value Like "FS*" And wsDL6850.Range("H" & i).Value = "TWD" Then
            Select Case wsDL6850.Range("N" & i).Value
                Case 0 To 10
                    copyCount0To10 = copyCount0To10 + 1
                    If copyCount0To10 > 10 Then
                        MsgBox "此期間流出之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow0TO10 & ":I" & destRow0TO10).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow0TO10 = destRow0TO10 + 1

                Case 11 To 30
                    copyCount11To30 = copyCount11To30 + 1
                    If copyCount11To30 > 10 Then
                        MsgBox "此期間流出之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow11TO30 & ":I" & destRow11TO30).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow11TO30 = destRow11TO30 + 1

                Case 31 To 90
                    copyCount31To90 = copyCount31To90 + 1
                    If copyCount31To90 > 10 Then
                        MsgBox "此期間流出之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow31TO90 & ":I" & destRow31TO90).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow31TO90 = destRow31TO90 + 1

                Case 91 To 180
                    copyCount91To180 = copyCount91To180 + 1
                    If copyCount91To180 > 10 Then
                        MsgBox "此期間流出之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow91TO180 & ":I" & destRow91TO180).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow91TO180 = destRow91TO180 + 1

                Case 181 To 365
                    copyCount181To365 = copyCount181To365 + 1
                    If copyCount181To365 > 10 Then
                        MsgBox "此期間流出之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow181TO365 & ":I" & destRow181TO365).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow181TO365 = destRow181TO365 + 1
            End Select
        End If
    Next i
    
    
    
    ' SPOT(FS) and InFlow TWD(colE)
    '起始貼入的目標列
    destRow0TO10 = 90
    destRow11TO30 = 100
    destRow31TO90 = 110
    destRow91TO180 = 120
    destRow181TO365 = 130

    '初始化計數變數
    copyCount0To10 = 0
    copyCount11To30 = 0
    copyCount31To90 = 0
    copyCount91To180 = 0
    copyCount181To365 = 0

    For i = 2 To rowCount
        If wsDL6850.Range("A" & i).Value Like "FS*" And wsDL6850.Range("E" & i).Value = "TWD" Then
            Select Case wsDL6850.Range("N" & i).Value
                Case 0 To 10
                    copyCount0To10 = copyCount0To10 + 1
                    If copyCount0To10 > 10 Then
                        MsgBox "此期間流入之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow0TO10 & ":T" & destRow0TO10).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow0TO10 = destRow0TO10 + 1

                Case 11 To 30
                    copyCount11To30 = copyCount11To30 + 1
                    If copyCount11To30 > 10 Then
                        MsgBox "此期間流入之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow11TO30 & ":T" & destRow11TO30).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow11TO30 = destRow11TO30 + 1

                Case 31 To 90
                    copyCount31To90 = copyCount31To90 + 1
                    If copyCount31To90 > 10 Then
                        MsgBox "此期間流入之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow31TO90 & ":T" & destRow31TO90).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow31TO90 = destRow31TO90 + 1

                Case 91 To 180
                    copyCount91To180 = copyCount91To180 + 1
                    If copyCount91To180 > 10 Then
                        MsgBox "此期間流入之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow91TO180 & ":T" & destRow91TO180).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow91TO180 = destRow91TO180 + 1

                Case 181 To 365
                    copyCount181To365 = copyCount181To365 + 1
                    If copyCount181To365 > 10 Then
                        MsgBox "此期間流入之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow181TO365 & ":T" & destRow181TO365).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow181TO365 = destRow181TO365 + 1
            End Select
        End If
    Next i
    











    'Case for over one year
    ' SWOP(SS or SF) and OutFlow TWD(colH)
    '起始貼入的目標列
    '初始化計數變數
    destRow366TO = 153
    copyCount366To = 0


    For i = 2 To rowCount
        If (wsDL6850.Range("A" & i).Value Like "SS*" Or wsDL6850.Range("A" & i).Value Like "SF*") And wsDL6850.Range("H" & i).Value = "TWD" Then
            Select Case wsDL6850.Range("N" & i).Value
                Case Is >= 366
                    copyCount366To = copyCount366To + 1
                    If copyCount366To > 10 Then
                        MsgBox "此期間流出之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow366TO & ":I" & destRow366TO).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow366TO = destRow366TO + 1
            End Select
        End If
    Next i
    
    
    
    ' SWOP(SS or SF) and InFlow TWD(colE)
    '起始貼入的目標列
    '初始化計數變數
    destRow366TO = 153
    copyCount366To = 0


    For i = 2 To rowCount
        If (wsDL6850.Range("A" & i).Value Like "SS*" Or wsDL6850.Range("A" & i).Value Like "SF*") And wsDL6850.Range("E" & i).Value = "TWD" Then
            Select Case wsDL6850.Range("N" & i).Value
                Case Is >= 366
                    copyCount366To = copyCount366To + 1
                    If copyCount366To > 10 Then
                        MsgBox "此期間流入之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow366TO & ":T" & destRow366TO).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow366TO = destRow366TO + 1
            End Select
        End If
    Next i
    
    
    
    ' SPOT(FS) and OutFlow TWD(colH)
    '起始貼入的目標列
    '初始化計數變數
    destRow366TO = 170
    copyCount366To = 0



    For i = 2 To rowCount
        If wsDL6850.Range("A" & i).Value Like "FS*" And wsDL6850.Range("H" & i).Value = "TWD" Then
            Select Case wsDL6850.Range("N" & i).Value
                Case Is >= 366
                    copyCount366To = copyCount366To + 1
                    If copyCount366To > 10 Then
                        MsgBox "此期間流出之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow366TO & ":I" & destRow366TO).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow366TO = destRow366TO + 1
            End Select
        End If
    Next i
    
    
    
    ' SPOT(FS) and InFlow TWD(colE)
    '起始貼入的目標列
    '初始化計數變數
    destRow366TO = 170
    copyCount366To = 0



    For i = 2 To rowCount
        If wsDL6850.Range("A" & i).Value Like "FS*" And wsDL6850.Range("E" & i).Value = "TWD" Then
            Select Case wsDL6850.Range("N" & i).Value
                Case Is >= 366
                    copyCount366To = copyCount366To + 1
                    If copyCount366To > 10 Then
                        MsgBox "此期間流入之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow366TO & ":T" & destRow366TO).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow366TO = destRow366TO + 1
            End Select
        End If
    Next i





    '完成
    MsgBox "完成"
End Sub
Sub ImportDL6850CSV()
    Dim wbImport As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim filePath As String
    Dim lastRow As Long

    ' 開啟檔案選擇視窗
    filePath = Application.GetOpenFilename("Excel 檔案 (*.xls), *.xls", , "請選擇 DL6850 Excel 檔")
    If filePath = "False" Then Exit Sub '使用者按取消

    ' 開啟選取的 CSV 檔（轉為 Excel 格式）
    Workbooks.Open Filename:=filePath
    Set wbImport = ActiveWorkbook
    Set wsSource = wbImport.Sheets(1)

    ' 指定貼上的目標工作表
    Set wsDest = ThisWorkbook.Sheets("DL6850原始資料")

    ' 找出來源的最後一列（避免多餘空白列）
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    ' 清除目標區原有資料（可視需求調整）
    wsDest.Range("B2:L" & wsDest.Rows.Count).ClearContents

    ' 貼上來源 A:L 到目標 B:L（從第2列開始貼）
    wsSource.Range("A1:L" & lastRow).Copy
    wsDest.Range("B1").PasteSpecial xlPasteValues

    ' 關閉 CSV 檔，不儲存
    Application.DisplayAlerts = False
    wbImport.Close SaveChanges:=False
    Application.DisplayAlerts = True

    MsgBox "DL6850 資料匯入完成！", vbInformation
End Sub




' ======================
F1_F2

Option Explicit

Sub MainSub()
    
    '清空 底稿_無NT、國外即期、國外換匯、國內即期、國內換匯資料
    ClearRange "底稿_無NT"
    ClearRange "國外即期"
    ClearRange "國外換匯"
    ClearRange "國內即期"
    ClearRange "國內換匯"
    

    '底稿_無NT
    Dim lastRow As Long
    Dim i As Long
    lastRow = Sheets("底稿_含NT").Cells(Rows.Count, 1).End(xlUp).Row
    Dim destinationRow As Long
    destinationRow = 2
    For i = 2 To lastRow
        If Sheets("底稿_含NT").Cells(i, 13).Value = False Then
            Sheets("底稿_含NT").Rows(i).Copy Destination:=Sheets("底稿_無NT").Rows(destinationRow)
            destinationRow = destinationRow + 1
        End If
    Next i

    '國外即期
    lastRow = Sheets("底稿_無NT").Cells(Rows.Count, 1).End(xlUp).Row
    destinationRow = 2
    For i = 2 To lastRow
        If Sheets("底稿_無NT").Cells(i, 1).Value = "FS" And Sheets("底稿_無NT").Cells(i, 11).Value = "國外" Then
            Sheets("底稿_無NT").Rows(i).Copy Destination:=Sheets("國外即期").Rows(destinationRow)
            destinationRow = destinationRow + 1
        End If
    Next i

    '國外換匯
    destinationRow = 2
    For i = 2 To lastRow
        If Sheets("底稿_無NT").Cells(i, 1).Value = "SS" And Sheets("底稿_無NT").Cells(i, 11).Value = "國外" Then
            Sheets("底稿_無NT").Rows(i).Copy Destination:=Sheets("國外換匯").Rows(destinationRow)
            destinationRow = destinationRow + 1
        End If
    Next i

    '國內即期
    destinationRow = 2
    For i = 2 To lastRow
        If Sheets("底稿_無NT").Cells(i, 1).Value = "FS" And Sheets("底稿_無NT").Cells(i, 11).Value = "國內" Then
            Sheets("底稿_無NT").Rows(i).Copy Destination:=Sheets("國內即期").Rows(destinationRow)
            destinationRow = destinationRow + 1
        End If
    Next i

    '國內換匯
    destinationRow = 2
    For i = 2 To lastRow
        If Sheets("底稿_無NT").Cells(i, 1).Value = "SS" And Sheets("底稿_無NT").Cells(i, 11).Value = "國內" Then
            Sheets("底稿_無NT").Rows(i).Copy Destination:=Sheets("國內換匯").Rows(destinationRow)
            destinationRow = destinationRow + 1
        End If
    Next i
    
    MsgBox "已完成"

End Sub


Sub ClearRange(sheetName As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rangeToClear As Range

    ' 定義要清空的工作表名稱
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' 取得最後一行和最後一列的位置
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' 定義要清空的範圍
    Set rangeToClear = ws.Range("A2:M" & lastRow)

    ' 清空範圍內的數值
    rangeToClear.ClearContents
End Sub


' ====================
FM2
要寫一段vba
1.將資料貼入 OBU_MM4901B分頁
2.以交易編號從A到Z排序，建立一個資料結構，當跑到第一個交易對象時，將該交易對象名稱及row儲存起來，如果有同樣交易對象出現，就在已經紀錄的交易對象新增那個row，如果是沒有的交易對象，則新增該交易對象及row儲存起來
3.將2資料結構中的交易對象在"金融機構代號對照表"分頁中A欄位逐一檢查，如果是在A欄位中的其中一個，新增一個歸類分類為DBU，如果是在B欄位中的其中一個，歸類分類為OBU，如果都沒找到，刪除該筆資料，將DBU及OBU分類儲存在資料結構中，以及c欄位的銀行代號也要儲存起來

4.
DBU 
INDEX1: ROW 3 TO 10
INDEX2: ROW 12 TO 19
INDEX3: ROW 21 TO 28
INDEX4: ROW 30 TO 37
INDEX5: ROW 39 TO 46

OBU
INDEX1: ROW 48 TO 55
INDEX2: ROW 57 TO 64
INDEX3: ROW 66 TO 73
INDEX4: ROW 75 TO 82
INDEX5: ROW 83 TO 90

上述3.中已經分配好的項目中，要逐一將結構中儲存的DBU及OBU，
依照所記載在資料結構中的資料，將分類屬於DBU的在 OBU_MM4901B 中紀錄的相關ROWS中的a:k欄位資料按照INDEX順序，一個INDEX代表一個交易對象，貼入DBU所列的相關ROW中
OBU亦同

5.另外按照上述方法，將記錄起來的銀行代號，逐一貼入FM2分頁C10往下的欄位中

' Chatgpt回答

下面我直接給你一個完整的 VBA 程式（單一 Sub），會依你描述的流程去處理。程式中我做了合理的假設並加了清楚註解、錯誤處理與執行摘要。請先把原始資料貼到工作表 OBU_MM4901B（假設標題列在第1列，資料從第2列開始），然後執行 FM2_ProcessOBU 即可。

    重要假設（程式以此運作；若你資料的欄位不是這樣，請告訴我我再調整）
    • OBU_MM4901B：資料範圍為 A:K（A=交易編號、B=交易對象名稱、其餘欄位照原樣）
    • 金融機構代號對照表：A欄放可判為 DBU 的交易對象名稱、B欄放可判為 OBU 的交易對象名稱、C欄為該筆的銀行代號（對應同列）
    • FM2：要把銀行代號寫到 C10 向下（依資料順序）
    • 每個交易對象 (counterparty) 的所有交易列會被視為該 counterparty 的“rows”，會在 INDEX 區塊中依序貼入（每個 INDEX 的高度為 8 列）
    • 若 DBU/OBU 的交易對象超過 5 個或某交易對象的筆數超過其 INDEX 的 8 列容量，會貼入可放的部份並在最後顯示警示（並用 Debug.Print 列出未放的筆數）
    
    ⸻
    
    
    Option Explicit
    
    Sub FM2_ProcessOBU()
        ' 主流程：依照使用者要求處理 OBU_MM4901B -> 分類 -> 填回區塊 -> 寫入 FM2 C10 開始的銀行代號
        ' 假設：資料原已貼到工作表 "OBU_MM4901B"，表頭在第1列，資料從第2列開始，資料欄位為 A:K (11欄)
        '         金融機構代號對照表: 工作表名 "金融機構代號對照表" ，A欄 DBU 名稱、B欄 OBU 名稱、C欄 銀行代號
        ' 使用說明：將原始資料貼到 OBU_MM4901B，執行此 Sub
        
        Dim wsData As Worksheet, wsMap As Worksheet, wsFM2 As Worksheet
        Dim lastRow As Long, dataRows As Long
        Dim arrData As Variant
        Dim dict As Object ' Scripting.Dictionary - key = partnerName, value = Collection of row-value-arrays
        Dim filteredDict As Object
        Dim i As Long, j As Long
        Dim partner As String
        Dim colRows As Collection
        Dim vRow() As Variant
        Dim findR As Range
        Dim key As Variant
        Dim DBUList As Collection, OBUList As Collection
        Dim DBU_bankCodes As Collection, OBU_bankCodes As Collection
        Dim DBUIndexTops As Variant, OBUIndexTops As Variant
        Dim idx As Long, pasteRow As Long
        Dim blockHeight As Long: blockHeight = 8 ' 每個INDEX的列數
        Dim maxIndexes As Long: maxIndexes = 5
        Dim msg As String
        Dim extraSkipped As Long
        
        On Error GoTo ErrHandler
        
        Set wsData = ThisWorkbook.Worksheets("OBU_MM4901B")
        Set wsMap = ThisWorkbook.Worksheets("金融機構代號對照表")
        Set wsFM2 = ThisWorkbook.Worksheets("FM2")
        
        ' 找出資料最後列（以 A 欄判斷）
        lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
        If lastRow < 2 Then
            MsgBox "OBU_MM4901B 找不到資料 (需至少有第2列開始的資料)。", vbExclamation
            Exit Sub
        End If
        
        dataRows = lastRow - 1
        ' 讀整個 A:K 資料到記憶體 (1-based array)
        arrData = wsData.Range("A2:K" & lastRow).Value
        
        ' 建 dictionary 儲存 partner -> collection(rows as variant arrays)
        Set dict = CreateObject("Scripting.Dictionary")
        dict.CompareMode = vbTextCompare ' 不分大小寫
        
        For i = 1 To UBound(arrData, 1)
            partner = Trim(CStr(arrData(i, 2))) ' 假設交易對象在 B 欄 (arrData col 2)
            If Len(partner) > 0 Then
                If Not dict.Exists(partner) Then
                    Set colRows = New Collection
                    dict.Add partner, colRows
                Else
                    Set colRows = dict(partner)
                End If
                ' 儲存該列的 A:K 值 (複製成一個一維陣列)
                ReDim vRow(1 To 11)
                For j = 1 To 11
                    vRow(j) = arrData(i, j)
                Next j
                colRows.Add vRow
            End If
        Next i
        
        ' 依照交易編號(A欄)排序：因為我們已讀入 arrData 並依序建立 dict（原始 arrData 是資料在工作表的順序）
        ' 如果你希望先依 A 欄排序再建立 dict，理想做法是把 arrData 先排序再走上面那段。以下為簡單作法：重新用 Worksheet.Sort 排序工作表的實際資料範圍 A:K
        With wsData
            .Range("A2:K" & lastRow).Sort Key1:=.Range("A2"), Order1:=xlAscending, Header:=xlNo
        End With
        ' 重新讀取 arrData 並重建 dict（確保 dict 的順序遵從交易編號 A->Z）
        arrData = wsData.Range("A2:K" & lastRow).Value
        Set dict = CreateObject("Scripting.Dictionary")
        dict.CompareMode = vbTextCompare
        For i = 1 To UBound(arrData, 1)
            partner = Trim(CStr(arrData(i, 2)))
            If Len(partner) > 0 Then
                If Not dict.Exists(partner) Then
                    Set colRows = New Collection
                    dict.Add partner, colRows
                Else
                    Set colRows = dict(partner)
                End If
                ReDim vRow(1 To 11)
                For j = 1 To 11
                    vRow(j) = arrData(i, j)
                Next j
                colRows.Add vRow
            End If
        Next i
        
        ' 檢查每個 partner 在 金融機構代號對照表 中是否為 A欄(DBU) 或 B欄(OBU)，
        ' 若找不到則刪除此 partner（從 dict 過濾掉）
        Set filteredDict = CreateObject("Scripting.Dictionary")
        filteredDict.CompareMode = vbTextCompare
        Set DBUList = New Collection
        Set OBUList = New Collection
        Set DBU_bankCodes = New Collection
        Set OBU_bankCodes = New Collection
        
        For Each key In dict.Keys
            partner = CStr(key)
            ' 先找 A 欄
            Set findR = wsMap.Columns("A").Find(What:=partner, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
            If Not findR Is Nothing Then
                ' DBU
                filteredDict.Add partner, dict(partner)
                DBUList.Add partner
                DBU_bankCodes.Add Trim(CStr(wsMap.Cells(findR.Row, "C").Value))
                GoTo NextKey
            End If
            ' 再找 B 欄
            Set findR = wsMap.Columns("B").Find(What:=partner, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
            If Not findR Is Nothing Then
                filteredDict.Add partner, dict(partner)
                OBUList.Add partner
                OBU_bankCodes.Add Trim(CStr(wsMap.Cells(findR.Row, "C").Value))
                GoTo NextKey
            End If
            ' 找不到 -> 不加入 filteredDict (會被視為刪除)
    NextKey:
        Next key
        
        ' 替換 dict 為 filteredDict（只剩找到的）
        Set dict = filteredDict
        
        ' 清空要貼入的目標區塊 (A:K 對每個 index 的範圍)
        DBUIndexTops = Array(3, 12, 21, 30, 39)   ' 對應你提供的 DBU index 起始列
        OBUIndexTops = Array(48, 57, 66, 75, 83)  ' OBU 起始列
        ' 清空這些區塊
        For idx = 0 To UBound(DBUIndexTops)
            wsData.Range(wsData.Cells(DBUIndexTops(idx), "A"), wsData.Cells(DBUIndexTops(idx) + blockHeight - 1, "K")).ClearContents
        Next idx
        For idx = 0 To UBound(OBUIndexTops)
            wsData.Range(wsData.Cells(OBUIndexTops(idx), "A"), wsData.Cells(OBUIndexTops(idx) + blockHeight - 1, "K")).ClearContents
        Next idx
        
        ' 把 DBU 的交易對象依序放入 DBU 的 index 區塊 (每個 index = 一個交易對象)
        extraSkipped = 0
        If DBUList.Count > 0 Then
            For idx = 1 To DBUList.Count
                If idx > maxIndexes Then
                    extraSkipped = extraSkipped + 1
                    Debug.Print "超過 DBU index 上限，跳過: " & DBUList(idx)
                Else
                    Dim topRow As Long
                    topRow = DBUIndexTops(idx - 1)
                    Set colRows = dict(DBUList(idx))
                    ' 將該交易對象的每列依序貼入 index block 的 A:K (最多 blockHeight 列)
                    For j = 1 To colRows.Count
                        If j > blockHeight Then
                            extraSkipped = extraSkipped + 1
                            Debug.Print "DBU (" & DBUList(idx) & ") 的第 " & j & " 列超過單一 index 的容量，已跳過。"
                        Else
                            ' 貼值
                            wsData.Range(wsData.Cells(topRow + j - 1, "A"), wsData.Cells(topRow + j - 1, "K")).Value = colRows(j)
                        End If
                    Next j
                End If
            Next idx
        End If
        
        ' 把 OBU 的交易對象依序放入 OBU 的 index 區塊
        If OBUList.Count > 0 Then
            For idx = 1 To OBUList.Count
                If idx > maxIndexes Then
                    extraSkipped = extraSkipped + 1
                    Debug.Print "超過 OBU index 上限，跳過: " & OBUList(idx)
                Else
                    topRow = OBUIndexTops(idx - 1)
                    Set colRows = dict(OBUList(idx))
                    For j = 1 To colRows.Count
                        If j > blockHeight Then
                            extraSkipped = extraSkipped + 1
                            Debug.Print "OBU (" & OBUList(idx) & ") 的第 " & j & " 列超過單一 index 的容量，已跳過。"
                        Else
                            wsData.Range(wsData.Cells(topRow + j - 1, "A"), wsData.Cells(topRow + j - 1, "K")).Value = colRows(j)
                        End If
                    Next j
                End If
            Next idx
        End If
        
        ' 把所有記錄到的銀行代號按順序貼入 FM2 C10 向下
        ' 先清除 C10 開始向下一些空間（例如清 200 行避免殘留）
        wsFM2.Range("C10:C500").ClearContents
        Dim outRow As Long: outRow = 10
        ' 按照最初的 encounter 順序（我們把 DBUList 與 OBUList 各自依序加入；但題意是“逐一貼入記錄起來的銀行代號”，
        ' 我保留先把 DBU 再把 OBU 寫入。如果你要「混合原始資料順序」請告訴我）
        For i = 1 To DBU_bankCodes.Count
            wsFM2.Cells(outRow, "C").Value = DBU_bankCodes(i)
            outRow = outRow + 1
        Next i
        For i = 1 To OBU_bankCodes.Count
            wsFM2.Cells(outRow, "C").Value = OBU_bankCodes(i)
            outRow = outRow + 1
        Next i
        
        ' 結果摘要
        msg = "處理完成。" & vbCrLf
        msg = msg & "原始資料列數: " & dataRows & vbCrLf
        msg = msg & "DBU 個數: " & DBUList.Count & vbCrLf
        msg = msg & "OBU 個數: " & OBUList.Count & vbCrLf
        If extraSkipped > 0 Then
            msg = msg & "有 " & extraSkipped & " 筆資料超過區塊容量或超過 index 上限，部分已被跳過(請看 Immediate/Debug 視窗)。"
        End If
        MsgBox msg, vbInformation, "FM2 處理結果"
        
        Exit Sub
    
    ErrHandler:
        MsgBox "發生錯誤: " & Err.Number & " - " & Err.Description, vbCritical
    End Sub
    
    
    ⸻
    
    如果你想我再幫你調整以下任一項，告訴我即可，我會改程式給你：
        1.	交易對象不是在 B 欄（請告訴我實際欄位，例如 C 欄或 D 欄）→ 我會改變 arrData(i, 2) 的索引。
        2.	想要 把 DBU 與 OBU 的銀行代號混合（按原始交易出現順序） 寫入 FM2!C10（目前程式是先寫 DBU 再寫 OBU）。
        3.	想要把超過容量（或超 index 數）的資料全部移到另一個工作表方便你檢視（現在程式只 Debug.Print 與顯示數量）。
        4.	想要把貼入 target 區塊的動作改為插入新列而不是覆寫（或要自動往下推移）。
    
    要不要我把上述任一個（或全部）改成你偏好的行為？

    ' =======================



' Question
根據我上述的資料

DBU 
INDEX1: ROW 3 TO 10
INDEX2: ROW 12 TO 19
INDEX3: ROW 21 TO 28
INDEX4: ROW 30 TO 37
INDEX5: ROW 39 TO 46

OBU
INDEX1: ROW 48 TO 55
INDEX2: ROW 57 TO 64
INDEX3: ROW 66 TO 73
INDEX4: ROW 75 TO 82
INDEX5: ROW 83 TO 90

我預留給DBU及OBU欄位的交易對手分別各有五個，每個交易對手最多8筆，如果有超過此DBU或OBU的交易對手數量，或是每個交易對手的交易筆數超過8筆，需要提供相關提醒，請幫我加入檢核這部分交易對手數量及各別交易筆數相關代碼，
請提供完整程式碼，並標記修改了哪邊

' Answer

太棒 — 我已把「檢核交易對手數量（>5）」與「每個交易對手交易筆數（>8）」的檢查完整加入程式，並會在執行結束時 顯示詳細警示（包含哪些交易對手超過 index 上限、哪些交易對手單筆超過 8 筆），且保留原本會跳過超出部分的行為（同時也會在 Immediate 視窗印出詳細 Debug 訊息）。
下面我把完整程式碼貼上，並在程式中用 ''' ### MODIFIED ### 與 ''' ### NOTE ### 標明我新增/修改的區段與說明，方便你快速找到變動處。

使用說明：把資料貼到 OBU_MM4901B（標題列為第1列，資料自第2列起，欄位 A:K），金融機構代號對照表 以 A 欄(DBU 名稱)、B 欄(OBU 名稱)、C 欄(銀行代號)，執行 FM2_ProcessOBU_WithChecks。

⸻


Option Explicit

Sub FM2_ProcessOBU_WithChecks()
    ' 主流程：依據使用者之規格處理 OBU_MM4901B -> 分類 -> 檢核 -> 填回區塊 -> 寫入 FM2 C10 開始的銀行代號
    ' 重要假設與欄位請參考先前說明 (交易對象在 B 欄，資料 A:K)
    
    Dim wsData As Worksheet, wsMap As Worksheet, wsFM2 As Worksheet
    Dim lastRow As Long, dataRows As Long
    Dim arrData As Variant
    Dim dict As Object ' partner -> Collection(rows as variant arrays)
    Dim i As Long, j As Long
    Dim partner As String
    Dim colRows As Collection
    Dim vRow() As Variant
    Dim findR As Range
    Dim key As Variant
    Dim DBUList As Collection, OBUList As Collection
    Dim DBU_bankCodes As Collection, OBU_bankCodes As Collection
    Dim DBUIndexTops As Variant, OBUIndexTops As Variant
    Dim idx As Long, pasteRow As Long
    Dim blockHeight As Long: blockHeight = 8 ' 每個INDEX的列數
    Dim maxIndexes As Long: maxIndexes = 5
    Dim msg As String
    Dim extraSkipped As Long
    
    ' === 以下為新增加用於檢核的變數 ===
    ''' ### MODIFIED ### 新增檢核用集合以記錄錯誤/超量資訊
    Dim overIndexDBU As Collection, overIndexOBU As Collection
    Dim overRowsList As Collection ' 每個超過 8 筆的紀錄 (字串描述)
    Dim partnerOrder As Collection ' 記錄處理順序（若需要混合順序寫FM2可用）
    ''' ### MODIFIED END ###
    
    On Error GoTo ErrHandler
    
    Set wsData = ThisWorkbook.Worksheets("OBU_MM4901B")
    Set wsMap = ThisWorkbook.Worksheets("金融機構代號對照表")
    Set wsFM2 = ThisWorkbook.Worksheets("FM2")
    
    ' 找出資料最後列（以 A 欄判斷）
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "OBU_MM4901B 找不到資料 (需至少有第2列開始的資料)。", vbExclamation
        Exit Sub
    End If
    
    dataRows = lastRow - 1
    ' 先用 Worksheet.Sort 依交易編號(A)排序，確保交易編號由 A->Z
    With wsData
        .Range("A2:K" & lastRow).Sort Key1:=.Range("A2"), Order1:=xlAscending, Header:=xlNo
    End With
    
    ' 讀整個 A:K 資料到記憶體
    arrData = wsData.Range("A2:K" & lastRow).Value
    
    ' 建 dictionary 儲存 partner -> collection(rows)
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare ' 不分大小寫
    
    ' 依排序後的 arrData 建立 dict（保持「第一次遇到 partner 的順序」）
    Set partnerOrder = New Collection
    For i = 1 To UBound(arrData, 1)
        partner = Trim(CStr(arrData(i, 2))) ' 假設交易對象在 B 欄
        If Len(partner) > 0 Then
            If Not dict.Exists(partner) Then
                Set colRows = New Collection
                dict.Add partner, colRows
                partnerOrder.Add partner ' 記錄遇到順序
            Else
                Set colRows = dict(partner)
            End If
            ' 儲存該列 A:K 值
            ReDim vRow(1 To 11)
            For j = 1 To 11
                vRow(j) = arrData(i, j)
            Next j
            colRows.Add vRow
        End If
    Next i
    
    ' 檢查 partner 是否在 金融機構代號對照表 裡面 (A->DBU, B->OBU)
    Set DBUList = New Collection
    Set OBUList = New Collection
    Set DBU_bankCodes = New Collection
    Set OBU_bankCodes = New Collection
    
    For Each key In dict.Keys
        partner = CStr(key)
        Set findR = wsMap.Columns("A").Find(What:=partner, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        If Not findR Is Nothing Then
            DBUList.Add partner
            DBU_bankCodes.Add Trim(CStr(wsMap.Cells(findR.Row, "C").Value))
            GoTo NextCheck
        End If
        Set findR = wsMap.Columns("B").Find(What:=partner, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        If Not findR Is Nothing Then
            OBUList.Add partner
            OBU_bankCodes.Add Trim(CStr(wsMap.Cells(findR.Row, "C").Value))
            GoTo NextCheck
        End If
        ' 找不到則不納入任何 List（相當於刪除該筆）
NextCheck:
    Next key
    
    ' === 新增檢核：檢查 DBU/OBU 交易對手總數是否超過 5 及每個交易對手筆數是否超過 8 ===
    ''' ### MODIFIED ### 初始化檢核集合
    Set overIndexDBU = New Collection
    Set overIndexOBU = New Collection
    Set overRowsList = New Collection
    ''' ### MODIFIED END ###
    
    ' 檢核 DBU 個數
    If DBUList.Count > maxIndexes Then
        ' 記錄超過的那些 partner (超出上限的那部分)
        For i = maxIndexes + 1 To DBUList.Count
            overIndexDBU.Add DBUList(i)
        Next i
    End If
    ' 檢核 OBU 個數
    If OBUList.Count > maxIndexes Then
        For i = maxIndexes + 1 To OBUList.Count
            overIndexOBU.Add OBUList(i)
        Next i
    End If
    
    ' 檢核每個 partner 的筆數是否超過 blockHeight (無論 DBU 或 OBU)
    Dim tmpCount As Long
    For i = 1 To DBUList.Count
        tmpCount = dict(DBUList(i)).Count
        If tmpCount > blockHeight Then
            overRowsList.Add "DBU | " & DBUList(i) & " : " & tmpCount & " 筆 (上限 " & blockHeight & " 筆)"
        End If
    Next i
    For i = 1 To OBUList.Count
        tmpCount = dict(OBUList(i)).Count
        If tmpCount > blockHeight Then
            overRowsList.Add "OBU | " & OBUList(i) & " : " & tmpCount & " 筆 (上限 " & blockHeight & " 筆)"
        End If
    Next i
    ''' ### MODIFIED ### 檢核結束 - 已把超量資訊記錄到 overIndexDBU / overIndexOBU / overRowsList
    ''' ### MODIFIED END ###
    
    ' 把 dict 中僅有的（也就是已在對照表找到的）分別依序貼回 DBU / OBU 區塊
    ' 我們需要以 DBUList 與 OBUList 的順序貼入。若 partner 在 dict 但未在 DBUList/OBUList（被排除）則不處理。
    
    DBUIndexTops = Array(3, 12, 21, 30, 39)   ' DBU index 起始列
    OBUIndexTops = Array(48, 57, 66, 75, 83)  ' OBU 起始列
    
    ' 清空目標區塊
    For idx = 0 To UBound(DBUIndexTops)
        wsData.Range(wsData.Cells(DBUIndexTops(idx), "A"), wsData.Cells(DBUIndexTops(idx) + blockHeight - 1, "K")).ClearContents
    Next idx
    For idx = 0 To UBound(OBUIndexTops)
        wsData.Range(wsData.Cells(OBUIndexTops(idx), "A"), wsData.Cells(OBUIndexTops(idx) + blockHeight - 1, "K")).ClearContents
    Next idx
    
    ' 貼入 DBU
    extraSkipped = 0
    If DBUList.Count > 0 Then
        For idx = 1 To DBUList.Count
            If idx > maxIndexes Then
                extraSkipped = extraSkipped + dict(DBUList(idx)).Count
                Debug.Print "超過 DBU index 上限，跳過整個交易對手: " & DBUList(idx)
            Else
                Dim topRow As Long
                topRow = DBUIndexTops(idx - 1)
                Set colRows = dict(DBUList(idx))
                For j = 1 To colRows.Count
                    If j > blockHeight Then
                        extraSkipped = extraSkipped + 1
                        Debug.Print "DBU (" & DBUList(idx) & ") 的第 " & j & " 列超過單一 index 的容量，已跳過。"
                    Else
                        wsData.Range(wsData.Cells(topRow + j - 1, "A"), wsData.Cells(topRow + j - 1, "K")).Value = colRows(j)
                    End If
                Next j
            End If
        Next idx
    End If
    
    ' 貼入 OBU
    If OBUList.Count > 0 Then
        For idx = 1 To OBUList.Count
            If idx > maxIndexes Then
                extraSkipped = extraSkipped + dict(OBUList(idx)).Count
                Debug.Print "超過 OBU index 上限，跳過整個交易對手: " & OBUList(idx)
            Else
                topRow = OBUIndexTops(idx - 1)
                Set colRows = dict(OBUList(idx))
                For j = 1 To colRows.Count
                    If j > blockHeight Then
                        extraSkipped = extraSkipped + 1
                        Debug.Print "OBU (" & OBUList(idx) & ") 的第 " & j & " 列超過單一 index 的容量，已跳過。"
                    Else
                        wsData.Range(wsData.Cells(topRow + j - 1, "A"), wsData.Cells(topRow + j - 1, "K")).Value = colRows(j)
                    End If
                Next j
            End If
        Next idx
    End If
    
    ' 把所有記錄到的銀行代號按順序貼入 FM2 C10 向下
    wsFM2.Range("C10:C500").ClearContents
    Dim outRow As Long: outRow = 10
    ' 這裡保留先寫 DBU 再寫 OBU，如需「混合原始出現順序」可改用 partnerOrder 及比對哪一類別來輸出
    For i = 1 To DBU_bankCodes.Count
        wsFM2.Cells(outRow, "C").Value = DBU_bankCodes(i)
        outRow = outRow + 1
    Next i
    For i = 1 To OBU_bankCodes.Count
        wsFM2.Cells(outRow, "C").Value = OBU_bankCodes(i)
        outRow = outRow + 1
    Next i
    
    ' 結果摘要與檢核訊息顯示
    msg = "處理完成。" & vbCrLf
    msg = msg & "原始資料列數: " & dataRows & vbCrLf
    msg = msg & "DBU 個數(已找到對應): " & DBUList.Count & vbCrLf
    msg = msg & "OBU 個數(已找到對應): " & OBUList.Count & vbCrLf
    If extraSkipped > 0 Then
        msg = msg & "注意：有 " & extraSkipped & " 筆資料因超出 INDEX/容量而被跳過(已寫入 Immediate 視窗)。" & vbCrLf
    End If
    
    ' === 新增：組合檢核訊息並顯示給使用者 ===
    ''' ### MODIFIED ### 若有超過 index 數量或超過 row 上限，彙總並顯示詳細清單
    Dim warnMsg As String
    warnMsg = ""
    If overIndexDBU.Count > 0 Then
        warnMsg = warnMsg & "DBU 超過 index 上限 (超過 5 個)：" & vbCrLf
        For i = 1 To overIndexDBU.Count
            warnMsg = warnMsg & "  - " & overIndexDBU(i) & vbCrLf
        Next i
    End If
    If overIndexOBU.Count > 0 Then
        warnMsg = warnMsg & "OBU 超過 index 上限 (超過 5 個)：" & vbCrLf
        For i = 1 To overIndexOBU.Count
            warnMsg = warnMsg & "  - " & overIndexOBU(i) & vbCrLf
        Next i
    End If
    If overRowsList.Count > 0 Then
        warnMsg = warnMsg & "交易筆數超過單一 index 上限 (每個交易對手上限 " & blockHeight & " 筆)：" & vbCrLf
        For i = 1 To overRowsList.Count
            warnMsg = warnMsg & "  - " & overRowsList(i) & vbCrLf
        Next i
    End If
    
    If Len(warnMsg) > 0 Then
        msg = msg & vbCrLf & "檢核警示（請注意）:" & vbCrLf & warnMsg
        ' 也在 Immediate 視窗印出更詳細 Debug
        Debug.Print "=== 檢核警示 ==="
        Debug.Print warnMsg
        MsgBox msg, vbExclamation, "FM2 處理結果（含檢核警示）"
    Else
        MsgBox msg, vbInformation, "FM2 處理結果"
    End If
    ''' ### MODIFIED END ###
    
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤: " & Err.Number & " - " & Err.Description, vbCritical
End Sub


⸻

我修改 / 新增的地方（重點標示）
	1.	Sub 名稱改為 FM2_ProcessOBU_WithChecks（完整程式請直接貼到模組中執行）。
	2.	在變數區新增：overIndexDBU, overIndexOBU, overRowsList, partnerOrder，用來記錄檢核結果與 partner 的出現順序。  ''' ### MODIFIED ### 在程式內有標示。
	3.	在建立 dict 時同時記錄原始「遇到順序」到 partnerOrder（若你要把 FM2 的銀行代號寫成原始出現順序可用此變數改寫輸出邏輯）。
	4.	新增檢核流程：
	•	檢查 DBU 與 OBU 各自的交易對手數量是否超過 maxIndexes（=5），若超過則把超出那部分的 partner 記錄到 overIndexDBU / overIndexOBU。
	•	檢查每個交易對手的交易筆數是否超過 blockHeight（=8），超過的則把描述加入 overRowsList。
以上區塊有標示 ''' ### MODIFIED ###。
	5.	將檢核結果整合到最後的 MsgBox 與 Immediate（Debug.Print）輸出，讓你能看到詳細清單（哪些 partner 超量、哪些 partner 筆數過多）。
	6.	保留原本「超出 capacity 或 index 上限會跳過」的行為，並會把跳過情況寫入 Immediate（也會在 summary 顯示跳過總數）。

⸻

如果你想要的進一步調整（我可以接著改）
	•	將被跳過（超過 index 或超過 8 筆）的那部分另存到另一個工作表，方便檢視／匯出。
	•	把 FM2 的銀行代號輸出改為原始出現順序混合 DBU/OBU（目前程式是先寫 DBU 再寫 OBU）。若要混合輸出，我會以 partnerOrder 去判斷每個 partner 屬於 DBU 或 OBU 再輸出銀行代號。
	•	若要在貼入時自動合併多於 8 筆的交易到下一個可用 index（例如把第1個 partner 前 8 筆放 index1，第9~16 筆放 index2），但題意好像每個 index 代表一個交易對手，所以我暫時以超量跳過處理；若你要改這個規則我也可以改。

要我把超出部分輸出到新工作表並把範例訊息顯示成 CSV 格式，或把銀行代號輸出順序改成原始出現順序嗎？或者你要我直接把這個程式放到你的活檔做測試（你可以貼 sample 資料）？


' ==============

FM10

Option Explicit

Sub CopyAndDeleteRows()
    Dim wsAC4603 As Worksheet
    Dim wsFM10 As Worksheet
    Dim n As Long
    Dim count As Long
    
    '<若欄位異動需更改項目>
    'AC4603檢核總欄位數，若AC4603欄位異動，確認所需欄位個數「更改下列count數值」
    count = 26

    ' 設定工作表名稱
    Set wsAC4603 = ThisWorkbook.Sheets("OBU_AC4603")
    Set wsFM10 = ThisWorkbook.Sheets("FM10底稿")

    ' 找到第n行的位置
    n = Application.Match("強制FVPL金融資產-公債-地方政府(外國)", wsAC4603.Range("A:A"), 0)

    ' 檢查是否找到了第n行
    If Not IsError(n) Then
        ' 檢查條件是否成立
        
        '---------------------------------------------
        '<若欄位異動需更改項目>
        '若欄位數異動，更改以下需檢核之欄位，欄位名稱需與報表名稱完全一致
        If wsAC4603.Range("A" & n + 1).Value = "強制FVPL金融資產-普通公司債(民營)(外國)" And _
           wsAC4603.Range("A" & n + 2).Value = "12005" And _
           wsAC4603.Range("A" & n + 3).Value = "強制FVPL金融資產評價調整-公債-地方-外國" And _
           wsAC4603.Range("A" & n + 4).Value = "強制FVPL金融資產評價調整-普通公司債(民營)(外國)" And _
           wsAC4603.Range("A" & n + 5).Value = "12007" And _
           wsAC4603.Range("A" & n + 6).Value = "FVOCI債務工具-公債-中央政府(外國)" And _
           wsAC4603.Range("A" & n + 7).Value = "FVOCI債務工具-普通公司債(公營)(外國)" And _
           wsAC4603.Range("A" & n + 8).Value = "FVOCI債務工具-普通公司債(民營)(外國)" And _
           wsAC4603.Range("A" & n + 9).Value = "FVOCI債務工具-金融債券-海外" And _
           wsAC4603.Range("A" & n + 10).Value = "12111" And _
           wsAC4603.Range("A" & n + 11).Value = "FVOCI債務工具評價調整-公債-中央政府(外國)" And _
           wsAC4603.Range("A" & n + 12).Value = "FVOCI債務工具評價調整-普通公司債(公營)(外國)" And _
           wsAC4603.Range("A" & n + 13).Value = "FVOCI債務工具評價調整-普通公司債(民營)(外國)" And _
           wsAC4603.Range("A" & n + 14).Value = "FVOCI債務工具評價調整-金融債券-海外" And _
           wsAC4603.Range("A" & n + 15).Value = "12113" And _
           wsAC4603.Range("A" & n + 16).Value = "AC債務工具投資-公債-中央政府(外國)" And _
           wsAC4603.Range("A" & n + 17).Value = "AC債務工具投資-普通公司債(民營)(外國)" And _
           wsAC4603.Range("A" & n + 18).Value = "AC債務工具投資-金融債券-海外" And _
           wsAC4603.Range("A" & n + 19).Value = "12201" And _
           wsAC4603.Range("A" & n + 20).Value = "累積減損-AC債務工具投資-公債-中央政府(外國)" And _
           wsAC4603.Range("A" & n + 21).Value = "累積減損-AC債務工具投資-普通公司(民營)(外國)" And _
           wsAC4603.Range("A" & n + 22).Value = "累積減損-AC債務工具投資-金融債券-海外" And _
           wsAC4603.Range("A" & n + 23).Value = "12203" And _
           wsAC4603.Range("A" & n + 24).Value = "拆放證券公司-OSU" And _
           wsAC4603.Range("A" & n + 25).Value = "15551" Then

            ' 刪除第n+count行至最後一行
            wsAC4603.Rows(n + count & ":" & wsAC4603.Rows.count).Delete

            ' 刪除第一行至第n-1行
            wsAC4603.Rows("1:" & n - 1).Delete
            
            '清除FM10底稿checkBox資料
            wsFM10.Range("A4:J" & (4 + count - 1)).ClearContents
            Application.CutCopyMode = False

            ' 複製AC4603數值內容到FM10底稿checkBox
            wsAC4603.Range("A1:J" & count).Copy
            wsFM10.Range("A4").Resize(count, 10).PasteSpecial Paste:=xlPasteValues
            
            MsgBox "完成"
            
            
        Else
            MsgBox "欄位有誤"
        End If
    Else
        MsgBox "找不到目標欄位 'FVOCI債務工具-公債-中央政府(外國)'"
    End If
End Sub



' ==================

FM11

Sub 匯入並篩選OBUAC5411B資料()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim importWB As Workbook
    Dim importFile As String
    Dim lastRow As Long, destRow As Long
    Dim cell As Range
    Dim keywords As Variant
    Dim i As Long
    Dim sumRange As Range

    ' 選取檔案
    importFile = Application.GetOpenFilename("Excel Files (*.xls;*.xlsx), *.xls;*.xlsx", , "請選取 OBU-AC5411B 檔案")
    If importFile = "False" Then Exit Sub ' 使用者取消

    ' 開啟來源檔案
    Set importWB = Workbooks.Open(importFile)
    
    ' 檢查是否存在名為 OBU-AC5411B 的分頁
    On Error Resume Next
    Set wsSource = importWB.Sheets(1)
    On Error GoTo 0
    If wsSource Is Nothing Then
        MsgBox "來源檔案中找不到分頁『OBU-AC5411B』", vbExclamation
        importWB.Close False
        Exit Sub
    End If

    ' 清除目前工作簿的 OBU-AC5411B 分頁舊資料（從第2列開始）
    Set wsDest = ThisWorkbook.Sheets("OBU-AC5411B")
    wsDest.Range("A2:Z" & wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row).ClearContents

    ' 複製來源檔案中第2列起資料貼到目前檔案中（保留標題列）
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    If lastRow >= 2 Then
        wsSource.Range("A2:Z" & lastRow).Copy Destination:=wsDest.Range("A2")
    End If
    
    ' 將 B 欄強制轉換成數值格式（避免 VLOOKUP 比對不到）
    With wsDest
        With .Range("B2:B" & .Cells(.Rows.Count, "B").End(xlUp).Row)
            .NumberFormat = "0"  ' 設定格式為一般數字
            .Value = .Value      ' 將文字轉換為純數值
        End With
    End With
    
    ' 關閉來源檔案
    importWB.Close False

    ' ---------- 以下是篩選與統計程式 ----------
    
    keywords = Array("FVPL", "FVOCI", "AC", "拆放證券公司息-OSU")

    ' 若目標工作表存在就刪除重建
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("OBU-AC5411B會科整理").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("OBU-AC5411B"))
    wsDest.Name = "OBU-AC5411B會科整理"

    ' 複製標題列
    ThisWorkbook.Sheets("OBU-AC5411B").Rows(1).Copy Destination:=wsDest.Rows(1)
    destRow = 2

    ' 遍歷 A 欄，找出符合關鍵字的列
    lastRow = ThisWorkbook.Sheets("OBU-AC5411B").Cells(ThisWorkbook.Sheets("OBU-AC5411B").Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        For Each keyword In keywords
            If InStr(ThisWorkbook.Sheets("OBU-AC5411B").Cells(i, "A").Value, keyword) > 0 Then
                ThisWorkbook.Sheets("OBU-AC5411B").Rows(i).Copy Destination:=wsDest.Rows(destRow)
                destRow = destRow + 1
                Exit For
            End If
        Next keyword
    Next i

    ' 總和 C 欄
    If destRow > 2 Then
        wsDest.Cells(destRow, "B").Value = "本月金額總和"
        Set sumRange = wsDest.Range("C2:C" & destRow - 1)
        wsDest.Cells(destRow, "C").Formula = "=SUM(" & sumRange.Address(False, False) & ")"
    
            ' 複製結果值到 FM11 計算1 的 G4 欄位（只取值，不取公式）
        ThisWorkbook.Sheets("FM11 計算1").Range("G4").Value = wsDest.Cells(destRow, "C").Value
    
    End If

    ' 自動欄寬
    wsDest.Columns.AutoFit

    MsgBox "匯入成功並完成篩選與總和計算！", vbInformation
    
    ThisWorkbook.Sheets("FM11 計算1").Activate
End Sub


' ===========

表41

Option Explicit

Sub SortAndCopyData()

    Dim wsDL9360 As Worksheet
    Dim wsTarget As Worksheet
    Dim baseDate As Date
    Dim exchangeRate As Double
    Dim lastRow As Long
    Dim n As Long
    Dim m As Long
    
    '設定工作表
    Set wsDL9360 = ThisWorkbook.Sheets("DL9360")
    Set wsTarget = ThisWorkbook.Sheets("底稿(扣掉TWD)")

    '彈出視窗，填寫基準日及美元兌換匯率
    baseDate = InputBox("請輸入基準日(日期格式：yyyy/mm/dd)", "基準日")
    exchangeRate = InputBox("請輸入美元兌換匯率", "美元兌換匯率")
    wsTarget.Range("C66").Value = baseDate
    wsTarget.Range("E66").Value = exchangeRate
    
    
    '刪除B欄位非日期格式及刪除國內交易對手(銀行國際代碼末4碼非TWTP)之交易
    lastRow = wsDL9360.Cells(wsDL9360.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = lastRow To 2 Step -1
        If Not IsDate(wsDL9360.Cells(i, "B").Value) Or Mid(wsDL9360.Cells(i, "E").Value, 5, 2) = "TW" Then
            wsDL9360.Rows(i).Delete
        End If
    Next i
    
    '清除 底稿(扣掉TWD) 工作表資料
    wsTarget.Range("A2:K31").ClearContents
    wsTarget.Range("A33:K62").ClearContents
    
    
    '重新取得欄位數
    lastRow = wsDL9360.Cells(wsDL9360.Rows.Count, "A").End(xlUp).Row

    ' 排序J欄位
    wsDL9360.Range("A1:K1").CurrentRegion.Sort Key1:=wsDL9360.Range("J2"), Order1:=xlAscending, Header:=xlYes


    ' 尋找小於零的最後一筆資料
    For n = 2 To lastRow
        If wsDL9360.Cells(n, "J").Value >= 0 Then
            m = lastRow - n + 1
            '當處分利益或損失任一交易筆數超過30筆時，中斷執行
            If n > 31 Or m > 30 Then
                MsgBox "筆數太多"
                Exit Sub
            Else
                Exit For
            End If
        End If
    Next n
    
    
    
    ' 複製資料至底稿工作表
    wsDL9360.Range("A2:K" & n - 1).Copy
    wsTarget.Range("A2").PasteSpecial Paste:=xlPasteValues

    wsDL9360.Range("A" & n & ":K" & lastRow).Copy
    wsTarget.Range("A33").PasteSpecial Paste:=xlPasteValues
  

    ' 清除剪貼板
    Application.CutCopyMode = False

    MsgBox "完成"
    
End Sub

' ===================



請問我要建立一個專案，需要將我之前在別支Excel寫的vba Sub全部打包起來，修改為只要執行一個Sub，
就可以依序將相關Sub依序都執行，相關Sub會修改原來VBA代碼所在Excel檔案中的分頁欄位資料，
請問這樣要怎麼做到，請給我一個簡化的範例，讓我知道怎麼運作


