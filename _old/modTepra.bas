'*****************************************************************************
' 備品管理ラベル印刷 サンプルプログラム for SPC10-API
' Copyright 2014 KING JIM CO.,LTD.
'*****************************************************************************

' モジュール名の指定
Attribute VB_Name =  "modTepra"


' Windows APIの関数を呼び出すVBA関数を定義
Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" _
                    (ByVal lpModuleName As String) As Long
                       
Declare PtrSafe Function GetProcAddress Lib "kernel32" _
                    (ByVal hModule As Long, _
                     ByVal lpProcName As String) As Long
                         
Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As Long

Declare PtrSafe Function IsWow64Process Lib "kernel32" _
                    (ByVal hProcess As Long, _
                     ByRef Wow64Process As Long) As Long

' データ開始行のオフセット
Public Const LINE_OFFSET As Integer = 19

' 対象データの最大行数
'Public Const MAX_LINE_COUNT As Integer = Range("D19").End(xlDown).Row
Public MAX_LINE_COUNT As Integer

' オプション有無文字列
Public const Flg_Yes as String = "する"
Public const Flg_No as String = "しない"

' シート名
Public Const strWsBihin As String = "備品管理台帳"
Public Const strWsAtesaki As String = "封筒宛先作成"

'エラーメッセージの定義
Public Const ERROR_MESSAGE_NO_PRINT_JOB = "印刷する項目が入力されていない､または印刷チェックマークが入っていないため､ラベルを印刷できません｡"
Public Const ERROR_MESSAGE_GET_TAPE_WIDTH = "テープ幅が取得できません。"
Public Const ERROR_MESSAGE_TPE_FILE_NOT_FOUND = "テープ幅に合ったレイアウトが存在しません。"
Public Const ERROR_MESSAGE_RUN_PRINT = """SPC10.exe""が指定した場所に存在しません。インストール先を確認してください。"

Public Const ERROR_MESSAGE_Job_Nothing = "印刷対象が選択されていません。"
Public Const ERROR_MESSAGE_Default_Template = "【テンプレート指定】で【指定しない】を選択した場合は、24mmか36mmのテープを本体に入れてください。"
Public Const ERROR_MESSAGE_Template_Nothing = "指定されたテンプレートが存在しません。"
Public Const ERROR_MESSAGE_Maisu_Nothing = "行目の【枚数】が指定されていません。"
Public Const ERROR_MESSAGE_Muki_Nothing = "行目の【向き】が指定されていません。"
	
	
'==============================================================================
' OSが64ビット環境かどうかを判別する関数の定義
'==============================================================================
Function IsWow64() As Boolean

'=============
'    Dim bIsWow64 As Long
'    Dim fnIsWow64Process As Long
'
'    ' 初期化
'    bIsWow64 = False
'    ' IsWow64Process関数が存在するかどうかを確認する
'    fnIsWow64Process = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")
'
'    If (0 <> fnIsWow64Process) Then
'        ' IsWow64ProcessはWindows XP SP2から導入された関数です
'        If 0 = IsWow64Process(GetCurrentProcess(), bIsWow64) Then
'            ' Windows XP の古いバージョンの場合はここに来ます
'        End If
'    End If
'=============

    Dim colItems As Object
    Dim itm As Object
    Dim ret As Boolean
     
    ret = False '初期化
    Set colItems = CreateObject("WbemScripting.SWbemLocator").ConnectServer.ExecQuery("Select * From Win32_OperatingSystem")
    For Each itm In colItems
        If InStr(itm.OSArchitecture, "64") Then
            ret = True
            Exit For
        End If
    Next

'    IsWow64 = (bIsWow64 <> 0)

    IsWow64 = ret

End Function

'==============================================================================
' テープ幅取得関数の定義
'==============================================================================
Function getTapeWidth(ByVal strFileName As String, ByRef strTapeType As String) As String
    
    Dim retValue As String
    
    ' ファイルを読み込むための配列
    Dim Arr()
    ReDim Preserve Arr(0)

    ' オブジェクトを作成
    Dim obj As Object
    Set obj = CreateObject("ADODB.Stream")

    ' オブジェクトに保存するデータの種類を文字列型に指定する
    obj.Type = adTypeText
    ' 文字列型のオブジェクトの文字コードを指定する
    obj.Charset = "UTF-16"

    ' オブジェクトのインスタンスを作成
    obj.Open

    ' ファイルからデータを読み込む
    obj.LoadFromFile (strFileName)

    ' 最終行までループする
    Do While Not obj.EOS
        ' 次の行を読み取る
        Arr(UBound(Arr)) = obj.ReadText(adReadLine)
        ReDim Preserve Arr(UBound(Arr) + 1)
    Loop

    ' オブジェクトを閉じる
    obj.Close

    ' メモリからオブジェクトを削除する
    Set obj = Nothing
    
    '-----------------------------------------------------------------------------
    ' テープ幅の取得
    '-----------------------------------------------------------------------------
    ' 読み込んだ1行目の文字列を分割（例：0x04 18mm）
    Dim strData As Variant
    strData = Split(Arr(0), " ")

    ' テープ幅を取得
    Dim strTapeWidth As String
    strTapeWidth = strData(0)

    ' テープ幅の設定
    Select Case strTapeWidth
        Case "0x00"
            retValue = "0"
        Case "0x01"
            retValue = "6"
        Case "0x02"
            retValue = "9"
        Case "0x03"
            retValue = "12"
        Case "0x04"
            retValue = "18"
        Case "0x05"
            retValue = "24"
        Case "0x06"
            retValue = "36"
        Case "0x0B"
            retValue = "4"
        Case "0x21"
            retValue = "50"
        Case "0x23"
            retValue = "100"
        Case "0xFF"
            retValue = ""
        Case Else
           retValue = ""
    End Select

    '-----------------------------------------------------------------------------
    ' テープ種類の取得
    '-----------------------------------------------------------------------------
    ' 読み込んだ2行目の文字列を分割（例：0x00 Standard tape）
    Dim strTypeData As Variant
    strTypeData = Split(Arr(1), " ")

    ' テープ種類を取得
    strTapeType = strTypeData(0)

    getTapeWidth = retValue
    
End Function

'==============================================================================
' オプション文字列生成関数の定義
'==============================================================================
Function createPrintOption(pathTempl As String, pathCsv As String, printNum As Integer, blnHalfcut As Boolean, blnConfirmTapeWidth As Boolean, strPrintLog As String, strTapeWidth As String) As String

    Dim comStrg   As String
    Dim retValue  As Double
    Dim strOption As String
    
    ' TPEファイルのフルパス名,CSVファイルのフルパス名,印刷部数
    strOption = pathTempl & "," & pathCsv & "," & printNum & ","
    
    ' テープ幅のファイル出力
    If (Len(strTapeWidth) > 0) Then
        strOption = strOption + "," + "/GT " + strTapeWidth
    End If
    
    ' カット設定
    If (blnHalfcut) Then
        strOption = strOption + "," + "/C -f -h"
    Else
        strOption = strOption + "," + "/C -f -hn"
    End If
    
    ' テープ幅確認メッセージのon/off設定
    If (blnConfirmTapeWidth) Then
        strOption = strOption + "," + "/TW -on"
    Else
        strOption = strOption + "," + "/TW -off"
    End If
    
    ' 印刷結果のファイル出力
    If (Len(strPrintLog) > 0) Then
        strOption = strOption + "," + "/L " + strPrintLog
    End If
   
    createPrintOption = strOption

End Function

'==============================================================================
' 印刷実行関数の定義
'==============================================================================
Function PrtSpc10Api(pathCmd As String, strOption As String, strPrinterName As String) As Double
    
    Dim comStrg  As String
    Dim retValue As Double
    
    ' 印刷コマンド
    If (Len(strPrinterName) > 0) Then
        ' /ptオプション
        comStrg = pathCmd & " " & "/pt " & Chr(34) & strOption & Chr(34) & " " & Chr(34) & strPrinterName & Chr(34)
    Else
        ' /pオプション
        comStrg = pathCmd & " " & "/p " & Chr(34) & strOption & Chr(34)
    End If
    
    On Error Resume Next
    
    retValue = Shell(comStrg, vbHide)
    
    PrtSpc10Api = retValue
    
End Function


Private sub ErrorConst()

	
end sub

