

Function GetDirFile(strFolderPath As String, Optional strExt As String = "") As String()
    '��1�����F�t�H���_�p�X
    '��2�����F�g���q(�����Ȃ��̓��C���h�J�[�h����)
    
    '�󂯎�����t�H���_���̃t�@�C�����擾���āA�z��ŕԂ�
    
    Dim cnt As Integer
    Dim buf As String
    cnt = 0
    
    Dim strFileName() As String
    'ReDim strFileName(cnt) As String
    
    Dim strSerchFile As String
    
    If strExt = "" Then
        strSerchFile = strFolderPath & "\*"
    Else
        strSerchFile = strFolderPath & "\*" & strExt
    End If
    
    buf = Dir(strSerchFile)
    Do While buf <> ""
        ReDim Preserve strFileName(cnt)
        strFileName(cnt) = buf  '�t�@�C������z��
        cnt = cnt + 1
        buf = Dir()     '���̃t�@�C�����擾
    Loop
    
    GetDirFile = strFileName
    
End Function

