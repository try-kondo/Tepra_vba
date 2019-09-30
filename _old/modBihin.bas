'*****************************************************************************
' 備品管理ラベル印刷 サンプルプログラム for SPC10-API
' Copyright 2014 KING JIM CO.,LTD.
'*****************************************************************************

'==============================================================================
' CSVファイルを作成し、印刷実行関数を呼び出す
'==============================================================================
Private Sub cmdPrint_Click()

    Dim strExePathName As String
    Dim strTextPathName As String
    Dim strCsvPathName As String
    Dim strPrintLogPathName As String
    Dim strTpePathName As String
    Dim strOption As String
    Dim dblRetValue As Double

    '対象のジョブ
    Dim strPrintjob() As String
    
    Dim strColName As String
    Dim strColDel As String
    
    MAX_LINE_COUNT = Range("D19").End(xlDown).Row
    
    '-----------------------------------------------------------------------------
    ' SPC10のEXEファイルをパス付きで指定する
    '-----------------------------------------------------------------------------
    If IsWow64() Then
        ' OSが64ビット環境
        strExePathName = "C:\Program Files (x86)\KING JIM\TEPRA SPC10\SPC10.exe"
    Else
        ' OSが32ビット環境
        strExePathName = "C:\Program Files\KING JIM\TEPRA SPC10\SPC10.exe"
    End If

    '-----------------------------------------------------------------------------
    ' テープ幅の出力ファイル、CSVファイル、印刷結果ファイル、TPEファイルの指定
    '-----------------------------------------------------------------------------
    strTextPathName = ThisWorkbook.Path & "\" & "TapeWidth.txt"       ' テープ幅の出力ファイルを指定する
    strCsvPathName = ThisWorkbook.Path & "\" & "data.csv"             ' CSVファイルを指定する
    strPrintLogPathName = ThisWorkbook.Path & "\" & "PrintResult.txt" ' 印刷結果ファイルを指定する
'    strTpePathName = ThisWorkbook.Path & "\bihin_18.tpe"              ' TPEファイルを指定する（デフォルト:18mm）
    ' TPEファイルを指定する（デフォルト:12mm - 1Line）
    strTpePathName = ThisWorkbook.Path & "\template\bihin_12_1line.tpe"
     
    '-----------------------------------------------------------------------------
    ' 設定情報の取得
    '-----------------------------------------------------------------------------
    ' ハーフカット設定
    Dim blnHalfcut As Boolean
    If OptionButton1.Value = True Then
        blnHalfcut = True
    ElseIf OptionButton2.Value = True Then
        blnHalfcut = False
    End If

    ' テープ幅確認メッセージ設定
    Dim blnConfirmTapeWidth As Boolean
    If chkTapeWidth.Value = True Then
        blnConfirmTapeWidth = True
    Else
        blnConfirmTapeWidth = False
    End If

    ' 印刷結果をファイルに出力する設定
    Dim strPrintLog As String
    If chkPrintLog.Value = True Then
        strPrintLog = strPrintLogPathName
    Else
        strPrintLog = ""
    End If

    '-----------------------------------------------------------------------------
    ' 印刷対象の確認
    '-----------------------------------------------------------------------------
'    If (getPrintJobCount() = 0) Then
'        MsgBox ERROR_MESSAGE_NO_PRINT_JOB
'        Exit Sub
'    End If

    strPrintjob() = getPrintJobCount()
    
    '-----------------------------------------------------------------------------
    ' テープ幅のファイル出力関数の呼び出し
    '-----------------------------------------------------------------------------
    strOption = createPrintOption(strTpePathName, strCsvPathName, 1, blnHalfcut, blnConfirmTapeWidth, strPrintLog, strTextPathName)
    dblRetValue = PrtSpc10Api(strExePathName, strOption, "")
    If (dblRetValue = 0) Then
        ' API実行エラー
        MsgBox ERROR_MESSAGE_RUN_PRINT
        Exit Sub
    End If

    '-----------------------------------------------------------------------------
    ' テープ幅の出力ファイルの存在確認
    '-----------------------------------------------------------------------------
    If Dir(strTextPathName) = "" Then
        ' テープ幅の出力ファイルが存在しない場合
        MsgBox ERROR_MESSAGE_GET_TAPE_WIDTH
        Exit Sub
    End If

    '-----------------------------------------------------------------------------
    ' TPEファイルの存在確認
    '-----------------------------------------------------------------------------
    Dim strTapeWidth As String
    Dim strTapeType As String
    
    ' テープ幅の出力ファイルからテープ幅（テープ種類）を取得
    strTapeType = ""
    strTapeWidth = getTapeWidth(strTextPathName, strTapeType)
    
    ' テープ幅の確認
    If StrComp(strTapeWidth, "0") = 0 Then
        ' テープ未装着の場合
        Exit Sub
    End If
    
    ' テープ種類の確認
    If StrComp(strTapeType, "0x00") Then
        ' Standard type以外の場合
        MsgBox ERROR_MESSAGE_TPE_FILE_NOT_FOUND
        Exit Sub
    End If

    ' TPEファイルを指定
    'カラム印刷有無チェック
    strColDel = ""
    If chkColDel.Value = True Then
        strColDel = "_col"
    End If
    
'    strTpePathName = ThisWorkbook.Path & "\bihin_" & strTapeWidth & ".tpe"
    strTpePathName = ThisWorkbook.Path & "\template\bihin_" & strTapeWidth & "_" & _
                     Application.WorksheetFunction.RoundUp(UBound(strPrintjob, 2) / 2, 0) & "line" & _
                     strColDel & ".tpe"
    
    If Dir(strTpePathName) = "" Then
        ' TPEファイルが存在しない場合
        MsgBox ERROR_MESSAGE_TPE_FILE_NOT_FOUND
        Exit Sub
    End If
    
    '-----------------------------------------------------------------------------
    ' CSVファイルの作成
    '-----------------------------------------------------------------------------
'    Dim csvStr1        As String
'    Dim csvStr2        As String
'    Dim csvStr3        As String
'    Dim csvStr4        As String
'    Dim csvQrStr       As String  ' QRコード
'    Dim csvStrAll      As String  ' 結合用
'    Dim strCompanyName As String  ' 会社名
'
'    Dim i As Integer
'    Dim chkValue(MAX_LINE_COUNT)
'    Dim fileNo As Integer

    On Error Resume Next
    
    ' 初期値設定
    fileNo = FreeFile
    
'    ' 会社名
'    strCompanyName = Range("E2").Value
    
    Open strCsvPathName For Output As #fileNo

    For i = LBound(strPrintjob, 1) To UBound(strPrintjob, 1)
        csvStrAll = ""
        For j = LBound(strPrintjob, 2) To UBound(strPrintjob, 2)
            csvStrAll = csvStrAll & strPrintjob(i, j)
            If j <> UBound(strPrintjob, 2) Then
                csvStrAll = csvStrAll & ","
            End If
        Next j
        
        Debug.Print csvStrAll
        'CSVファイル出力
        Print #fileNo, csvStrAll
    Next i
    
'    For i = 1 To MAX_LINE_COUNT
'        chkValue(i) = ActiveSheet.OLEObjects("CheckBox" & i).Object.Value
'        If chkValue(i) Then
'            ' チェックがされている場合
'            csvStr1 = Range("B" & (i + LINE_OFFSET)).Value
'            csvStr2 = Range("C" & (i + LINE_OFFSET)).Value
'            csvStr3 = Range("D" & (i + LINE_OFFSET)).Value
'            csvStr4 = Range("E" & (i + LINE_OFFSET)).Value
'            ' データの存在確認
'            If (Len(csvStr1) = 0 And Len(csvStr2) = 0 And Len(csvStr3) = 0 And Len(csvStr4) = 0) Then
'                ' 空データの行は、チェックされていても無視する
'            Else
'                ' QRコード用文字列
'                csvQrStr = Chr(34) & _
'                           csvStr1 & "," & _
'                           csvStr2 & "," & _
'                           csvStr3 & "," & _
'                           csvStr4 & _
'                           Chr(34)
'
'                ' 統合
'                csvStrAll = strCompanyName & "," & _
'                            csvStr1 & "," & _
'                            csvStr2 & "," & _
'                            csvStr3 & "," & _
'                            csvStr4 & "," & _
'                            csvQrStr
'
'                ' CSVファイル出力
'                Print #fileNo, csvStrAll
'            End If
'        End If
'    Next i
    
    Close #fileNo

    '-----------------------------------------------------------------------------
    ' 印刷関数の呼び出し
    '-----------------------------------------------------------------------------
    ' 印刷実行
    strOption = createPrintOption(strTpePathName, strCsvPathName, 1, blnHalfcut, blnConfirmTapeWidth, strPrintLog, "")
    dblRetValue = PrtSpc10Api(strExePathName, strOption, "")
    If (dblRetValue = 0) Then
        ' API実行エラー
        MsgBox ERROR_MESSAGE_RUN_PRINT
    End If

End Sub

'==============================================================================
' 印刷ジョブ数を取得する
'==============================================================================
Private Function getPrintJobCount() As String()

'    Dim csvStr1 As String
'    Dim csvStr2 As String
'    Dim csvStr3 As String
'    Dim csvStr4 As String
'    Dim intPrintJob As Integer ' 印刷JOB数
    Dim i As Integer
'    Dim chkValue(MAX_LINE_COUNT)
    
    Dim strPrintjob() As String     'CSV書き出し用
    Dim intPrintCol() As Integer    '印刷対象のカラムを格納
    Dim intMaxCol As Integer        '表の最大列
    intMaxCol = Range("D19").End(xlToRight).Column
    
    Dim intPrintRow() As Integer
    
    Dim cnt As Integer
    
    Dim j As Integer
    
    '対象のカラムの検索
    cnt = 0
    ReDim intPrintCol(intMaxCol)
    For i = 4 To intMaxCol
        If Cells(LINE_OFFSET - 1, i).Value = "○" Then
            intPrintCol(cnt) = i
            cnt = cnt + 1
        End If
    Next i
    
    'カラムが何も選択されていないとき
    If cnt = 0 Then
        MsgBox "印刷対象の列が選択されていません"
        End
    End If
    
    ReDim Preserve intPrintCol(cnt - 1)
    '対象行の検索
    cnt = 0
    ReDim intPrintRow(MAX_LINE_COUNT)
    For i = 20 To MAX_LINE_COUNT
        If Cells(i, 3).Value = "○" Then
            intPrintRow(cnt) = i
            cnt = cnt + 1
        End If
    Next i
    
    '行が何も選択されていないとき
    If cnt = 0 Then
        MsgBox "印刷対象の行が選択されていません"
        End
    End If
    
    ReDim Preserve intPrintRow(cnt - 1)
    
    'CSV作成用
    ReDim strPrintjob(UBound(intPrintRow), UBound(intPrintCol) * 2 + 1) 'カラム名も書き込むから列だけ２倍する
    For i = 0 To UBound(intPrintRow)
        cnt = 0
        For j = 0 To UBound(strPrintjob, 2) Step 2
            'カラム名書き込み
            strPrintjob(i, j) = Cells(LINE_OFFSET, intPrintCol(cnt)).Value
            '中身書き込み
            strPrintjob(i, j + 1) = Cells(intPrintRow(i), intPrintCol(cnt)).Value
            
            Debug.Print strPrintjob(i, j) & " , " & strPrintjob(i, j + 1)
            cnt = cnt + 1
        Next j
        
        Debug.Print "-------------"
    Next i
    
    Debug.Print "========================================================"
    
    getPrintJobCount = strPrintjob
    
        
'    ' 初期値設定
'    intPrintJob = 0
'
'    For i = 1 To MAX_LINE_COUNT
'        chkValue(i) = ActiveSheet.OLEObjects("CheckBox" & i).Object.Value
'        If chkValue(i) Then
'            ' チェックがされている場合
'            csvStr1 = Range("B" & (i + LINE_OFFSET)).Value
'            csvStr2 = Range("C" & (i + LINE_OFFSET)).Value
'            csvStr3 = Range("D" & (i + LINE_OFFSET)).Value
'            csvStr4 = Range("E" & (i + LINE_OFFSET)).Value
'            ' データの存在確認
'            If (Len(csvStr1) = 0 And Len(csvStr2) = 0 And Len(csvStr3) = 0 And Len(csvStr4) = 0) Then
'                ' 空データの行は、チェックされていても無視する
'            Else
'                intPrintJob = intPrintJob + 1
'            End If
'        End If
'    Next i
    
'    getPrintJobCount = intPrintJob
    
End Function

