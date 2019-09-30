'*****************************************************************************
' 備品管理ラベル印刷 サンプルプログラム for SPC10-API
' Copyright 2014 KING JIM CO.,LTD.
'*****************************************************************************

' モジュール名の指定
Attribute VB_Name =  "modBihin"

'==============================================================================
' CSVファイルを作成し、印刷実行関数を呼び出す
'==============================================================================
Public Sub Bihin_main()

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
    
    Dim WsOption As Worksheet
    Set WsOption = Worksheets("option")
    
    Dim WsList As Worksheet
    Set WsList = Worksheets("List")
    
    '備品管理台帳
    Dim WsBihin As Worksheet
    Set WsBihin = Worksheets(strWsBihin)
   	
    MAX_LINE_COUNT = WsBihin.Range("D19").End(xlDown).Row
    
    Dim strHalfcut as String
    Dim strConfirmTapeWidth as String
    Dim strOptPrintLog as String
    Dim strOptCol as String
    strHalfcut = WsBihin.Cells(WsOption.Range("C15").Value, WsOption.Range("C16").Value ).Value
    strConfirmTapeWidth = WsBihin.Cells(WsOption.Range("C17").Value, WsOption.Range("C18").Value ).Value
    strOptPrintLog = WsBihin.Cells(WsOption.Range("C19").Value, WsOption.Range("C20").Value ).Value
    strOptCol = WsBihin.Cells(WsOption.Range("C21").Value, WsOption.Range("C22").Value ).Value
    
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
	if Flg_Yes = strHalfcut Then
		blnHalfcut = True
	Else
		blnHalfcut = False
	End if
    
    ' テープ幅確認メッセージ設定
    Dim blnConfirmTapeWidth As Boolean
    if Flg_Yes = strConfirmTapeWidth Then
        blnConfirmTapeWidth = True
    Else
        blnConfirmTapeWidth = False
    End If

    ' 印刷結果をファイルに出力する設定
    Dim strPrintLog As String
    if Flg_Yes = strOptPrintLog Then
        strPrintLog = strPrintLogPathName
    Else
        strPrintLog = ""
    End If
    
    '-----------------------------------------------------------------------------
    ' 印刷対象の確認
    '-----------------------------------------------------------------------------
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
    'If WsBihin.chkColDel.Value = True Then
    if Flg_Yes = strOptCol Then
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

    Dim i As Integer
    
    Dim strPrintjob() As String     'CSV書き出し用
    Dim intPrintCol() As Integer    '印刷対象のカラムを格納
    Dim intMaxCol As Integer        '表の最大列
    intMaxCol = Range("D19").End(xlToRight).Column
    
    Dim intPrintRow() As Integer
    
    Dim cnt As Integer
    
    Dim j As Integer
    
    '備品管理台帳
    Dim WsBihin As Worksheet
    Set WsBihin = Worksheets(strWsBihin)
    
    '対象のカラムの検索
    cnt = 0
    ReDim intPrintCol(intMaxCol)
    For i = 4 To intMaxCol
        If WsBihin.Cells(LINE_OFFSET - 1, i).Value = "○" Then
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
        If WsBihin.Cells(i, 3).Value = "○" Then
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
            strPrintjob(i, j) = WsBihin.Cells(LINE_OFFSET, intPrintCol(cnt)).Value
            '中身書き込み
            strPrintjob(i, j + 1) = WsBihin.Cells(intPrintRow(i), intPrintCol(cnt)).Value
            
            cnt = cnt + 1
        Next j
        
    Next i
    
    getPrintJobCount = strPrintjob
    
End Function

