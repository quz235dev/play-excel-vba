Attribute VB_Name = "lib"
' UTF-8エンコードのファイルを読み込み、行の配列として返す関数
Function ReadUTF8File(filePath As String) As String()
    Dim fileContent As String
    Dim lines() As String
    
    ' ファイルの存在確認
    If Dir(filePath) = "" Then
        Err.Raise 53, "ReadUTF8File", "ファイルが見つかりません: " & filePath
        Exit Function
    End If
    
    ' ADODB.Streamを使ってUTF-8ファイルを読み込む
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    
    On Error GoTo ErrorHandler
    
    With stream
        .Charset = "UTF-8"
        .Type = 2 ' テキストモード
        .Open
        .LoadFromFile filePath
        
        ' BOMをスキップ
        If AscW(Left(.ReadText(3), 1)) = 239 Then
            .Position = 0
            .Charset = "UTF-8"
            .Position = 3 ' BOMをスキップ
        Else
            .Position = 0
        End If
        
        fileContent = .ReadText(-1) ' ファイル全体を読み込む
        .Close
    End With
    
    ' 行に分割（CRLF改行コードに対応）
    lines = Split(fileContent, vbCrLf)
    
    ReadUTF8File = lines
    Exit Function
    
ErrorHandler:
    If Not stream Is Nothing Then
        If stream.State = 1 Then ' 開いている場合
            stream.Close
        End If
    End If
    
    Err.Raise Err.Number, "ReadUTF8File", "ファイルの読み込みに失敗しました: " & Err.Description
End Function
