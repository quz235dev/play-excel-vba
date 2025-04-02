Attribute VB_Name = "Module2"
Sub ボタン2_Click()
    Dim filePath As String
    filePath = Application.GetOpenFilename("TSV Files (*.tsv),*.tsv", , "TSVファイルを選択してください")
    
    If filePath = "False" Then Exit Sub  ' ユーザーがキャンセルした場合
    
    Dim deviceData As Variant
    deviceData = ParseTSVFile(filePath)
    
    If Not IsArray(deviceData) Then
        MsgBox "ファイルの解析に失敗しました。", vbExclamation
        Exit Sub
    End If
    
    MsgBox "完了"

End Sub
