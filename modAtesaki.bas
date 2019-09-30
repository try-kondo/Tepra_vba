
Dim strPrintjob() As String
Dim strPrintOption() As String

Dim intTarRow(2) As Integer            '印刷対象表の最小行,最大行
Dim intTarCol(2) As Integer            '印刷対象表の最小列,最大列
Dim intOptRow(2) As Integer
Dim intOptCol(2) As Integer
    
Public Sub Atesaki_main()
    
        
    Dim WsOption As Worksheet
    Set WsOption = Worksheets("option")
    
    Dim strExePathName As String        'SPC10のEXEファイルパスを格納
    Dim strTextPathName As String
    Dim strCsvPathName As String
    Dim strPrintLogPathName As String
    Dim strTpePathName As String
    Dim strOption As String
    Dim dblRetValue As Double
    
    Dim csvStrAll As String
    
    Dim PrintLastFlag As Boolean
    
    '印刷対象範囲格納
    intTarRow(0) = WsOption.Range("D3").Value
    intTarCol(0) = WsOption.Range("D4").Value
    intTarRow(1) = Cells(intTarRow(0), intTarCol(0)).End(xlDown).Row
    intTarCol(1) = WsOption.Range("D6").Value
    
    intOptRow(0) = WsOption.Range("D7").Value
    intOptCol(0) = WsOption.Range("D8").Value
    intOptRow(1) = intTarRow(1)
    intOptCol(1) = WsOption.Range("D10").Value
    
    '代入時に増設していくからとりあえず１で宣言
    ReDim strPrintjob(0, intTarCol(1) - intTarCol(0))
    ReDim strPrintOption(0, intOptCol(1) - intOptCol(0))
    
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
    
    '-----------------------------------------------------------------------------
    ' 印刷対象の確認
    '-----------------------------------------------------------------------------
    DoEvents
    Call getPrintJobCount(strTapeWidth)
    
    '-----------------------------------------------------------------------------
    ' CSVファイルの作成
    '-----------------------------------------------------------------------------
    
    On Error Resume Next
    
    ' 初期値設定
    fileNo = FreeFile
    
    For i = LBound(strPrintjob, 1) To UBound(strPrintjob, 1)
        DoEvents
        'CSVファイルオープン
        Open strCsvPathName For Output As #fileNo
        '書き込み用変数初期化
        csvStrAll = ""
        'カラムくっつけ処理
        For j = LBound(strPrintjob, 2) To UBound(strPrintjob, 2)
            csvStrAll = csvStrAll & strPrintjob(i, j)
            If j <> UBound(strPrintjob, 2) Then     '最終カラムにはカンマはつけない
                csvStrAll = csvStrAll & ","
            End If
        Next j
        
        'CSVファイル出力
        Print #fileNo, csvStrAll
        
        '最終行チェック
        PrintLastFlag = False
        If i = UBound(strPrintjob, 1) Then
            PrintLastFlag = True
        End If
        
        '最終行チェック
        If PrintLastFlag = False Then
            '次の行が同じテンプレートかチェック
            If strPrintOption(i, 0) = strPrintOption(i + 1, 0) Then
                '同じ
                'Stop
            Else
                '違ければ、ファイルを閉じてテプラ出力
                Close #fileNo
                
                strTpePathName = strPrintOption(i, 0)
                
                '-----------------------------------------------------------------------------
                ' 印刷関数の呼び出し    ここの関数化めんどい
                '-----------------------------------------------------------------------------
                ' 印刷実行
                strOption = createPrintOption(strTpePathName, strCsvPathName, 1, blnHalfcut, blnConfirmTapeWidth, strPrintLog, "")
                dblRetValue = PrtSpc10Api(strExePathName, strOption, "")
                If (dblRetValue = 0) Then
                    ' API実行エラー
                    MsgBox ERROR_MESSAGE_RUN_PRINT
                End If
            End If
        Else
            Close #fileNo
                
            strTpePathName = strPrintOption(i, 0)
                
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
        End If
        
    Next i

End Sub

'==============================================================================
' 印刷ジョブ数を取得する
'==============================================================================
Private Function getPrintJobCount(strTapeWidth As String)

    Dim i As Integer
    Dim j As Integer
    
    Dim cntCol As Integer
    Dim cntRow As Integer
    Dim temp As Integer
    temp = 0
    
    Dim strTpePath As String
    strTpePath = ThisWorkbook.Path & "\template\"
    
    Dim intTarNum() As Integer
    Dim intTempRow() As Integer
    
    Dim strTarDirection As String      'テプラ向き
    
    '印刷枚数チェック
    cntRow = 0
    For i = intTarRow(0) To intTarRow(1)
        '印刷対象チェック
        If Cells(i, intTarCol(0) - 1) = "○" Then
            '枚数チェック
            'ReDim用
            '枚数が空白だったら終了
            If Cells(i, intOptCol(1)).Value = "" Or IsNumeric(Cells(i, intOptCol(1)).Value) = False Then
                MsgBox i & " " & ERROR_MESSAGE_Maisu_Nothing
                End
            End If
            temp = temp + Cells(i, intOptCol(1)).Value
            
            '後の繰り返し用
            ReDim Preserve intTarNum(cntRow)
            intTarNum(cntRow) = Cells(i, intOptCol(1)).Value
            
            '○の行を記憶
            ReDim Preserve intTempRow(cntRow)
            intTempRow(cntRow) = i
            
            cntRow = cntRow + 1
            
        End If
    Next i
    
    '印刷が選択されているか確認
    'されていなければ、エラーメッセージ表示後終了
    If cntRow = 0 Then
        MsgBox ERROR_MESSAGE_Job_Nothing
        End
    End If
    
    ReDim strPrintjob(temp - 1, UBound(strPrintjob, 2))
    ReDim strPrintOption(temp - 1, UBound(strPrintOption, 2))
    
    '対象の行を検索
    cntRow = 0
    For i = LBound(intTempRow) To UBound(intTempRow)    '行繰り返し
        For k = 0 To intTarNum(i) - 1                   '印刷枚数分だけ繰り返し
            cntCol = 0
            For j = intTarCol(0) To intTarCol(1)  '列繰り返し
                '印刷対象の格納
                strPrintjob(cntRow, cntCol) = Cells(intTempRow(i), j)
                cntCol = cntCol + 1
            Next j
            '対象のテンプレート格納
            '「指定しない」を選択した場合
            If Cells(intTempRow(i), intOptCol(0)) = "指定しない" Then
                '「縦」「横」チェック　それ以外はエラー
                If Cells(intTempRow(i), intOptCol(0) + 1) = "縦" Then
                    strTarDirection = "_tate"
                ElseIf Cells(intTempRow(i), intOptCol(0) + 1) = "横" Then
                    strTarDirection = "_yoko"
                Else
                    '向きが選択されていない場合は終了
                    MsgBox intTempRow(i) & " " & ERROR_MESSAGE_Muki_Nothing
                    End
                End If
                strPrintOption(cntRow, 0) = strTpePath & "atesaki_" & strTapeWidth & strTarDirection & ".tpe"
            'テンプレートを選択している場合
            Else
                strPrintOption(cntRow, 0) = strTpePath & Cells(intTempRow(i), intOptCol(0))
            End If
            'テンプレート存在チェックをしたい
            '指定しないの場合、テンプレートとテプラサイズが違う場合、エラーメッセージ
            If Dir(strPrintOption(cntRow, 0)) = "" Then
                If Cells(intTempRow(i), intOptCol(0)) = "指定しない" Then
                    MsgBox ERROR_MESSAGE_Default_Template & vbCrLf & vbCrLf & _
                           "本体に搭載されているテープ ： " & strTapeWidth & " mm"
                    End
                End If
                MsgBox ERROR_MESSAGE_Template_Nothing
                End
            End If
            
            cntRow = cntRow + 1
        Next k
    Next i
    
End Function

Public Sub Atesaki_Template()
    '================================================
    '封筒宛先作成シートのテンプレート入力規則を更新
    '================================================

    '変数宣言
    Dim WsOption As Worksheet
    Set WsOption = Worksheets("option")
    Dim WsList As Worksheet
    Set WsList = Worksheets("List")
    
    Dim strTempFile() As String         'テンプレートファイル名を格納
    Dim strTempPath As String           'テンプレートのパスを格納
    strTempPath = ThisWorkbook.Path & "\template"
    Dim strExt As String
    strExt = ".tpe"
    
    Dim intListRow(1) As Integer
    Dim intListCol(1) As Integer
    'リスト検索範囲
    intListRow(0) = WsOption.Range("D11").Value
    intListCol(0) = WsOption.Range("D12").Value
    intListRow(1) = WsOption.Range("D13").Value
    intListCol(1) = WsOption.Range("D14").Value
    
    Dim i As Integer
    Dim cnt As Integer
    
    'Listシートの封筒宛先列をクリアする
    '一番上は「指定しない」だからそれはクリアしない
    For i = intListRow(0) + 1 To intListRow(1)
        WsList.Cells(i, intListCol(0)).Clear
    Next i
    
    'フォルダ内けんさくけんさく
    'テンプレートフォルダ内の「.tpe」ファイルを取得する
    strTempFile = GetDirFile(strTempPath, strExt)
    
    'Listシートに書き出し
    '3行目から書き出し
    cnt = 3
    For i = LBound(strTempFile) To UBound(strTempFile)
        WsList.Cells(cnt, intListCol(0)).Value = strTempFile(i)
        cnt = cnt + 1
    Next i
    
    MsgBox "テンプレートファイルを更新しました。"
    
End Sub
