Attribute VB_Name = "perseTsv"


' UTF-8�G���R�[�h��TSV�t�@�C�����p�[�X���ăf�[�^��z��Ƃ��ĕԂ��֐�
Function ParseTSVFile(filePath As String) As Variant
    Dim deviceData() As Variant  ' �߂�l�p�̔z��
    Dim deviceCount As Integer
    
    On Error GoTo ErrorHandler
    
    ' UTF-8�t�@�C����ǂݍ��݁A�s�̔z��Ƃ��Ď擾
    Dim lines() As String
    lines = ReadUTF8File(filePath)
    
    ' �f�[�^��ێ�����z��
    deviceCount = 0
    ReDim deviceData(1 To 100) ' �����T�C�Y�i��Ń��T�C�Y�j
    
    Dim currentDevice As Variant
    Dim currentPhaseIndex As Integer
    Dim headers As Variant
    Dim phaseDataArray As Variant
    Dim phaseDict As Object
    
    currentPhaseIndex = 0
    
    ' �L����blockType�̃��X�g���`
    Dim validBlockTypes As Variant
    validBlockTypes = GetBlockTypes()
    
    Dim i As Integer
    For i = 0 To UBound(lines)
        Dim line As String
        line = Trim(lines(i))
        
        ' ��s�̓X�L�b�v (�u���b�N��؂�)
        If line = "" Then
            currentPhaseIndex = 0
            GoTo ContinueLoop
        End If
        
        ' �f�o�C�X�u���b�N���ʍs
        If InStr(line, vbTab) > 0 And currentPhaseIndex = 0 Then
            ' �f�o�C�X�u���b�N�������
            Dim parts As Variant
            parts = Split(line, vbTab)
            
            ' blockType���L�����X�g�Ɋ܂܂�Ă��邩�m�F
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
                ' �V�����f�o�C�X�f�[�^������
                deviceCount = deviceCount + 1
                If deviceCount > UBound(deviceData) Then
                    ReDim Preserve deviceData(1 To UBound(deviceData) * 2)
                End If
                
                ' �f�o�C�X�f�[�^�\����������
                ' �z��̊e�v�f�ɂ͈ȉ��̏�񂪊i�[�����:
                ' 0:blockType - �f�o�C�X�u���b�N�̎�ށi��: "upw-pump"�j
                ' 1:blockName - �f�o�C�X�u���b�N�̖��O�i��: "�T�u�|���v"�j
                ' 2:headers   - �t�B�[���h���̔z��i��: ["name", "value", "unit"]�j
                ' 3:phases    - 4�̃t�F�[�Y�f�[�^�����z��i1-4�j
                '               �e�t�F�[�Y��Dictionary�I�u�W�F�N�g�ŁA
                '               �t�B�[���h�����L�[�ɒl���i�[�i��: phases(1)("value") = "85"�j
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
        
        ' �w�b�_�[�s�̏���
        If currentPhaseIndex = 0 And deviceCount > 0 Then
            ' �w�b�_�[�����
            headers = Split(line, vbTab)
            deviceData(deviceCount)(2) = headers
            
            ' �t�F�[�Y�f�[�^��������
            ReDim phaseDataArray(1 To 4) ' 4�̃t�F�[�Y
            deviceData(deviceCount)(3) = phaseDataArray
            
            currentPhaseIndex = 1
            GoTo ContinueLoop
        End If
        
        ' �f�[�^�s�̏���
        If currentPhaseIndex >= 1 And currentPhaseIndex <= 4 And deviceCount > 0 Then
            ' �s�f�[�^�����
            Dim values As Variant
            values = Split(line, vbTab)
            
            ' �w�b�_�[�ɑΉ�����f�[�^�������`���Ŋi�[
            Set phaseDict = CreateObject("Scripting.Dictionary")
            
            Dim j As Integer
            For j = 0 To UBound(headers)
                If j <= UBound(values) Then
                    phaseDict.Add headers(j), values(j)
                Else
                    phaseDict.Add headers(j), ""
                End If
            Next j
            
            ' �t�F�[�Y�f�[�^��ۑ�
            Set deviceData(deviceCount)(3)(currentPhaseIndex) = phaseDict
            
            currentPhaseIndex = currentPhaseIndex + 1
        End If
        
ContinueLoop:
    Next i
    
    ' �z��𐳊m�ȃT�C�Y�Ƀ��T�C�Y
    If deviceCount < UBound(deviceData) Then
        ReDim Preserve deviceData(1 To deviceCount)
    End If
    
    ParseTSVFile = deviceData
    Exit Function
    
ErrorHandler:
    Dim errMsg As String
    errMsg = "�G���[���������܂���: " & Err.Description
    Debug.Print errMsg
    MsgBox errMsg, vbExclamation, "�G���["
    
    ' ��̔z���Ԃ�
    ReDim deviceData(0 To 0)
    ParseTSVFile = deviceData
End Function

