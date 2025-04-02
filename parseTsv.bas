Attribute VB_Name = "perseTsv"


' UTF-8エンコードのTSVファイルをパースしてデータを配列として返す関数
Function ParseTSVFile(filePath As String) As Variant
    Dim deviceData() As Variant  ' 戻り値用の配列
    Dim deviceCount As Integer
    
    On Error GoTo ErrorHandler
    
    ' UTF-8ファイルを読み込み、行の配列として取得
    Dim lines() As String
    lines = ReadUTF8File(filePath)
    
    ' データを保持する配列
    deviceCount = 0
    ReDim deviceData(1 To 100) ' 初期サイズ（後でリサイズ）
    
    Dim currentDevice As Variant
    Dim currentPhaseIndex As Integer
    Dim headers As Variant
    Dim phaseDataArray As Variant
    Dim phaseDict As Object
    
    currentPhaseIndex = 0
    
    ' 有効なblockTypeのリストを定義
    Dim validBlockTypes As Variant
    validBlockTypes = GetBlockTypes()
    
    Dim i As Integer
    For i = 0 To UBound(lines)
        Dim line As String
        line = Trim(lines(i))
        
        ' 空行はスキップ (ブロック区切り)
        If line = "" Then
            currentPhaseIndex = 0
            GoTo ContinueLoop
        End If
        
        ' デバイスブロック識別行
        If InStr(line, vbTab) > 0 And currentPhaseIndex = 0 Then
            ' デバイスブロック情報を解析
            Dim parts As Variant
            parts = Split(line, vbTab)
            
            ' blockTypeが有効リストに含まれているか確認
            Dim isValidBlockType As Boolean
            isValidBlockType = False
            
            Dim k As Integer
            For k = 0 To UBound(validBlockTypes)
                If parts(0) = validBlockTypes(k) Then
                    isValidBlockType = True
                    Exit For
                End If
            Next k
            
            If isValidBlockType Then
                ' 新しいデバイスデータを準備
                deviceCount = deviceCount + 1
                If deviceCount > UBound(deviceData) Then
                    ReDim Preserve deviceData(1 To UBound(deviceData) * 2)
                End If
                
                ' デバイスデータ構造を初期化
                ' 配列の各要素には以下の情報が格納される:
                ' 0:blockType - デバイスブロックの種類（例: "upw-pump"）
                ' 1:blockName - デバイスブロックの名前（例: "サブポンプ"）
                ' 2:headers   - フィールド名の配列（例: ["name", "value", "unit"]）
                ' 3:phases    - 4つのフェーズデータを持つ配列（1-4）
                '               各フェーズはDictionaryオブジェクトで、
                '               フィールド名をキーに値を格納（例: phases(1)("value") = "85"）
                ReDim currentDevice(3)
                
                currentDevice(0) = parts(0) ' blockType
                If UBound(parts) > 0 Then
                    currentDevice(1) = parts(1) ' blockName
                Else
                    currentDevice(1) = ""
                End If
                
                deviceData(deviceCount) = currentDevice
                currentPhaseIndex = 0
                GoTo ContinueLoop
            End If
            
        End If
        
        ' ヘッダー行の処理
        If currentPhaseIndex = 0 And deviceCount > 0 Then
            ' ヘッダーを解析
            headers = Split(line, vbTab)
            deviceData(deviceCount)(2) = headers
            
            ' フェーズデータを初期化
            ReDim phaseDataArray(1 To 4) ' 4つのフェーズ
            deviceData(deviceCount)(3) = phaseDataArray
            
            currentPhaseIndex = 1
            GoTo ContinueLoop
        End If
        
        ' データ行の処理
        If currentPhaseIndex >= 1 And currentPhaseIndex <= 4 And deviceCount > 0 Then
            ' 行データを解析
            Dim values As Variant
            values = Split(line, vbTab)
            
            ' ヘッダーに対応するデータを辞書形式で格納
            Set phaseDict = CreateObject("Scripting.Dictionary")
            
            Dim j As Integer
            For j = 0 To UBound(headers)
                If j <= UBound(values) Then
                    phaseDict.Add headers(j), values(j)
                Else
                    phaseDict.Add headers(j), ""
                End If
            Next j
            
            ' フェーズデータを保存
            Set deviceData(deviceCount)(3)(currentPhaseIndex) = phaseDict
            
            currentPhaseIndex = currentPhaseIndex + 1
        End If
        
ContinueLoop:
    Next i
    
    ' 配列を正確なサイズにリサイズ
    If deviceCount < UBound(deviceData) Then
        ReDim Preserve deviceData(1 To deviceCount)
    End If
    
    ParseTSVFile = deviceData
    Exit Function
    
ErrorHandler:
    Dim errMsg As String
    errMsg = "エラーが発生しました: " & Err.Description
    Debug.Print errMsg
    MsgBox errMsg, vbExclamation, "エラー"
    
    ' 空の配列を返す
    ReDim deviceData(0 To 0)
    ParseTSVFile = deviceData
End Function

