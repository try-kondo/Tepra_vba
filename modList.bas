

Function GetDirFile(strFolderPath As String, Optional strExt As String = "") As String()
    '第1引数：フォルダパス
    '第2引数：拡張子(引数なしはワイルドカード扱い)
    
    '受け取ったフォルダ内のファイルを取得して、配列で返す
    
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
        strFileName(cnt) = buf  'ファイル名を配列
        cnt = cnt + 1
        buf = Dir()     '次のファイルを取得
    Loop
    
    GetDirFile = strFileName
    
End Function

