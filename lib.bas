Attribute VB_Name = "lib"
' UTF-8�G���R�[�h�̃t�@�C����ǂݍ��݁A�s�̔z��Ƃ��ĕԂ��֐�
Function ReadUTF8File(filePath As String) As String()
    Dim fileContent As String
    Dim lines() As String
    
    ' �t�@�C���̑��݊m�F
    If Dir(filePath) = "" Then
        Err.Raise 53, "ReadUTF8File", "�t�@�C����������܂���: " & filePath
        Exit Function
    End If
    
    ' ADODB.Stream���g����UTF-8�t�@�C����ǂݍ���
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    
    On Error GoTo ErrorHandler
    
    With stream
        .Charset = "UTF-8"
        .Type = 2 ' �e�L�X�g���[�h
        .Open
        .LoadFromFile filePath
        
        ' BOM���X�L�b�v
        If AscW(Left(.ReadText(3), 1)) = 239 Then
            .Position = 0
            .Charset = "UTF-8"
            .Position = 3 ' BOM���X�L�b�v
        Else
            .Position = 0
        End If
        
        fileContent = .ReadText(-1) ' �t�@�C���S�̂�ǂݍ���
        .Close
    End With
    
    ' �s�ɕ����iCRLF���s�R�[�h�ɑΉ��j
    lines = Split(fileContent, vbCrLf)
    
    ReadUTF8File = lines
    Exit Function
    
ErrorHandler:
    If Not stream Is Nothing Then
        If stream.State = 1 Then ' �J���Ă���ꍇ
            stream.Close
        End If
    End If
    
    Err.Raise Err.Number, "ReadUTF8File", "�t�@�C���̓ǂݍ��݂Ɏ��s���܂���: " & Err.Description
End Function
