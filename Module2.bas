Attribute VB_Name = "Module2"
Sub �{�^��2_Click()
    Dim filePath As String
    filePath = Application.GetOpenFilename("TSV Files (*.tsv),*.tsv", , "TSV�t�@�C����I�����Ă�������")
    
    If filePath = "False" Then Exit Sub  ' ���[�U�[���L�����Z�������ꍇ
    
    Dim deviceData As Variant
    deviceData = ParseTSVFile(filePath)
    
    If Not IsArray(deviceData) Then
        MsgBox "�t�@�C���̉�͂Ɏ��s���܂����B", vbExclamation
        Exit Sub
    End If
    
    MsgBox "����"

End Sub
