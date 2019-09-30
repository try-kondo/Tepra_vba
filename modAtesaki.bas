
' ���W���[�����̎w��
Attribute VB_Name =  "modAtesaki"

Dim strPrintjob() As String
Dim strPrintOption() As String

Dim intTarRow(2) As Integer            '����Ώە\�̍ŏ��s,�ő�s
Dim intTarCol(2) As Integer            '����Ώە\�̍ŏ���,�ő��
Dim intOptRow(2) As Integer
Dim intOptCol(2) As Integer
    
Public Sub Atesaki_main()
    
    Dim WsOption As Worksheet
    Set WsOption = Worksheets("option")
    
    Dim WsList As Worksheet
    Set WsList = Worksheets("List")
    
    '��������쐬s
    Dim WsAtesaki As Worksheet
    Set WsAtesaki = Worksheets(strWsAtesaki)
    
    Dim strExePathName As String        'SPC10��EXE�t�@�C���p�X���i�[
    Dim strTextPathName As String
    Dim strCsvPathName As String
    Dim strPrintLogPathName As String
    Dim strTpePathName As String
    Dim strOption As String
    Dim dblRetValue As Double
    
    Dim csvStrAll As String
    
    Dim PrintLastFlag As Boolean
    
    '����Ώ۔͈͊i�[
    intTarRow(0) = WsOption.Range("D3").Value
    intTarCol(0) = WsOption.Range("D4").Value
    if WsAtesaki.Cells(intTarRow(0)+ 1, intTarCol(0)) = "" then
    	intTarRow(1) = 16
    Else
    	intTarRow(1) = WsAtesaki.Cells(intTarRow(0), intTarCol(0)).End(xlDown).Row
    End if
    intTarCol(1) = WsOption.Range("D6").Value
    
    intOptRow(0) = WsOption.Range("D7").Value
    intOptCol(0) = WsOption.Range("D8").Value
    intOptRow(1) = intTarRow(1)
    intOptCol(1) = WsOption.Range("D10").Value
    
    '������ɑ��݂��Ă�������Ƃ肠�����P�Ő錾
    ReDim strPrintjob(0, intTarCol(1) - intTarCol(0))
    ReDim strPrintOption(0, intOptCol(1) - intOptCol(0))
    
    Dim strHalfcut as String
    Dim strConfirmTapeWidth as String
    Dim strOptPrintLog as String
    Dim strOptCol as String
    strHalfcut = WsAtesaki.Cells(WsOption.Range("D15").Value, WsOption.Range("D16").Value ).Value
    strConfirmTapeWidth = WsAtesaki.Cells(WsOption.Range("D17").Value, WsOption.Range("D18").Value ).Value
    strOptPrintLog = WsAtesaki.Cells(WsOption.Range("D19").Value, WsOption.Range("D20").Value ).Value
    strOptCol = WsAtesaki.Cells(WsOption.Range("D21").Value, WsOption.Range("D22").Value ).Value
    
    '-----------------------------------------------------------------------------
    ' SPC10��EXE�t�@�C�����p�X�t���Ŏw�肷��
    '-----------------------------------------------------------------------------
    If IsWow64() Then
        ' OS��64�r�b�g��
        strExePathName = "C:\Program Files (x86)\KING JIM\TEPRA SPC10\SPC10.exe"
    Else
        ' OS��32�r�b�g��
        strExePathName = "C:\Program Files\KING JIM\TEPRA SPC10\SPC10.exe"
    End If
    
    '-----------------------------------------------------------------------------
    ' �e�[�v���̏o�̓t�@�C���ACSV�t�@�C���A������ʃt�@�C���ATPE�t�@�C���̎w��
    '-----------------------------------------------------------------------------
    strTextPathName = ThisWorkbook.Path & "\" & "TapeWidth.txt"       ' �e�[�v���̏o�̓t�@�C�����w�肷��
    strCsvPathName = ThisWorkbook.Path & "\" & "data.csv"             ' CSV�t�@�C�����w�肷��
    strPrintLogPathName = ThisWorkbook.Path & "\" & "PrintResult.txt" ' ������ʃt�@�C�����w�肷��
    strTpePathName = ThisWorkbook.Path & "\template\bihin_12_1line.tpe"
    
    '-----------------------------------------------------------------------------
    ' �ݒ���̎擾
    '-----------------------------------------------------------------------------
    ' �n�[�t�J�b�g�ݒ�
    Dim blnHalfcut As Boolean
	if Flg_Yes = strHalfcut Then
		blnHalfcut = True
	Else
		blnHalfcut = False
	End if
    
    ' �e�[�v���m�F���b�Z�[�W�ݒ�
    Dim blnConfirmTapeWidth As Boolean
    if Flg_Yes = strConfirmTapeWidth Then
        blnConfirmTapeWidth = True
    Else
        blnConfirmTapeWidth = False
    End If

    ' ������ʂ��t�@�C���ɏo�͂���ݒ�
    Dim strPrintLog As String
    if Flg_Yes = strOptPrintLog Then
        strPrintLog = strPrintLogPathName
    Else
        strPrintLog = ""
    End If
    
    '-----------------------------------------------------------------------------
    ' �e�[�v���̃t�@�C���o�͊֐��̌Ăяo��
    '-----------------------------------------------------------------------------
    strOption = createPrintOption(strTpePathName, strCsvPathName, 1, blnHalfcut, blnConfirmTapeWidth, strPrintLog, strTextPathName)
    dblRetValue = PrtSpc10Api(strExePathName, strOption, "")
    If (dblRetValue = 0) Then
        ' API���s�G���[
        MsgBox ERROR_MESSAGE_RUN_PRINT
        Exit Sub
    End If

    '-----------------------------------------------------------------------------
    ' �e�[�v���̏o�̓t�@�C���̑��݊m�F
    '-----------------------------------------------------------------------------
    If Dir(strTextPathName) = "" Then
        ' �e�[�v���̏o�̓t�@�C�������݂��Ȃ��ꍇ
        MsgBox ERROR_MESSAGE_GET_TAPE_WIDTH
        Exit Sub
    End If
    
    '-----------------------------------------------------------------------------
    ' TPE�t�@�C���̑��݊m�F
    '-----------------------------------------------------------------------------
    Dim strTapeWidth As String
    Dim strTapeType As String
    
    ' �e�[�v���̏o�̓t�@�C������e�[�v���i�e�[�v��ށj���擾
    strTapeType = ""
    strTapeWidth = getTapeWidth(strTextPathName, strTapeType)
    
    ' �e�[�v���̊m�F
    If StrComp(strTapeWidth, "0") = 0 Then
        ' �e�[�v�������̏ꍇ
        Exit Sub
    End If
    
    ' �e�[�v��ނ̊m�F
    If StrComp(strTapeType, "0x00") Then
        ' Standard type�ȊO�̏ꍇ
        MsgBox ERROR_MESSAGE_TPE_FILE_NOT_FOUND
        Exit Sub
    End If
    
    '-----------------------------------------------------------------------------
    ' ����Ώۂ̊m�F
    '-----------------------------------------------------------------------------
    DoEvents
    Call getPrintJobCount(strTapeWidth)
    
    '-----------------------------------------------------------------------------
    ' CSV�t�@�C���̍쐬
    '-----------------------------------------------------------------------------
    
    On Error Resume Next
    
    ' �����l�ݒ�
    fileNo = FreeFile
    
    For i = LBound(strPrintjob, 1) To UBound(strPrintjob, 1)
        DoEvents
        'CSV�t�@�C���I�[�v��
        Open strCsvPathName For Output As #fileNo
        '�������ݗp�ϐ�������
        csvStrAll = ""
        '�J��������������
        For j = LBound(strPrintjob, 2) To UBound(strPrintjob, 2)
            csvStrAll = csvStrAll & strPrintjob(i, j)
            If j <> UBound(strPrintjob, 2) Then     '�ŏI�J�����ɂ̓J���}�͂��Ȃ�
                csvStrAll = csvStrAll & ","
            End If
        Next j
        
        'CSV�t�@�C���o��
        Print #fileNo, csvStrAll
        
        '�ŏI�s�`�F�b�N
        PrintLastFlag = False
        If i = UBound(strPrintjob, 1) Then
            PrintLastFlag = True
        End If
        
        '�ŏI�s�`�F�b�N
        If PrintLastFlag = False Then
            '���̍s�������e���v���[�g���`�F�b�N
            If strPrintOption(i, 0) = strPrintOption(i + 1, 0) Then
                '����
                'Stop
            Else
                '�Ⴏ��΁A�t�@�C������ăe�v���o��
                Close #fileNo
                
                strTpePathName = strPrintOption(i, 0)
                
                '-----------------------------------------------------------------------------
                ' ����֐��̌Ăяo��    �����̊֐����߂�ǂ�
                '-----------------------------------------------------------------------------
                ' ������s
                strOption = createPrintOption(strTpePathName, strCsvPathName, 1, blnHalfcut, blnConfirmTapeWidth, strPrintLog, "")
                dblRetValue = PrtSpc10Api(strExePathName, strOption, "")
                If (dblRetValue = 0) Then
                    ' API���s�G���[
                    MsgBox ERROR_MESSAGE_RUN_PRINT
                End If
            End If
        Else
            Close #fileNo
                
            strTpePathName = strPrintOption(i, 0)
                
            '-----------------------------------------------------------------------------
            ' ����֐��̌Ăяo��
            '-----------------------------------------------------------------------------
            ' ������s
            strOption = createPrintOption(strTpePathName, strCsvPathName, 1, blnHalfcut, blnConfirmTapeWidth, strPrintLog, "")
            dblRetValue = PrtSpc10Api(strExePathName, strOption, "")
            If (dblRetValue = 0) Then
                ' API���s�G���[
                MsgBox ERROR_MESSAGE_RUN_PRINT
            End If
        End If
        
    Next i

End Sub

'==============================================================================
' ����W���u�����擾����
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
    
    Dim strTarDirection As String      '�e�v������
    
    '��������쐬
    Dim WsAtesaki As Worksheet
    Set WsAtesaki = Worksheets(strWsAtesaki)
    
    '��������`�F�b�N
    cntRow = 0
    For i = intTarRow(0) To intTarRow(1)
        '����Ώۃ`�F�b�N
        If WsAtesaki.Cells(i, intTarCol(0) - 1) = "��" Then
            '�����`�F�b�N
            'ReDim�p
            '�������󔒂�������I��
            If WsAtesaki.Cells(i, intOptCol(1)).Value = "" Or _
               IsNumeric(WsAtesaki.Cells(i, intOptCol(1)).Value) = False Then
                MsgBox i & " " & ERROR_MESSAGE_Maisu_Nothing
                End
            End If
            temp = temp + WsAtesaki.Cells(i, intOptCol(1)).Value
            
            '��̌J��Ԃ��p
            ReDim Preserve intTarNum(cntRow)
            intTarNum(cntRow) = WsAtesaki.Cells(i, intOptCol(1)).Value
            
            '���̍s���L��
            ReDim Preserve intTempRow(cntRow)
            intTempRow(cntRow) = i
            
            cntRow = cntRow + 1
            
        End If
    Next i
    
    '������I������Ă��邩�m�F
    '����Ă��Ȃ���΁A�G���[���b�Z�[�W�\����I��
    If cntRow = 0 Then
        MsgBox ERROR_MESSAGE_Job_Nothing
        End
    End If
    
    ReDim strPrintjob(temp - 1, UBound(strPrintjob, 2))
    ReDim strPrintOption(temp - 1, UBound(strPrintOption, 2))
    
    '�Ώۂ̍s������
    cntRow = 0
    For i = LBound(intTempRow) To UBound(intTempRow)    '�s�J��Ԃ�
        For k = 0 To intTarNum(i) - 1                   '��������������J��Ԃ�
            cntCol = 0
            For j = intTarCol(0) To intTarCol(1)  '��J��Ԃ�
                '����Ώۂ̊i�[
                strPrintjob(cntRow, cntCol) = WsAtesaki.Cells(intTempRow(i), j)
                cntCol = cntCol + 1
            Next j
            '�Ώۂ̃e���v���[�g�i�[
            '�u�w�肵�Ȃ��v��I�������ꍇ
            If WsAtesaki.Cells(intTempRow(i), intOptCol(0)) = "�w�肵�Ȃ�" Then
                '�u�c�v�u���v�`�F�b�N�@����ȊO�̓G���[
                If WsAtesaki.Cells(intTempRow(i), intOptCol(0) + 1) = "�c" Then
                    strTarDirection = "_tate"
                ElseIf WsAtesaki.Cells(intTempRow(i), intOptCol(0) + 1) = "��" Then
                    strTarDirection = "_yoko"
                Else
                    '�������I������Ă��Ȃ��ꍇ�͏I��
                    MsgBox intTempRow(i) & " " & ERROR_MESSAGE_Muki_Nothing
                    End
                End If
                strPrintOption(cntRow, 0) = strTpePath & "atesaki_" & strTapeWidth & strTarDirection & ".tpe"
            '�e���v���[�g��I�����Ă���ꍇ
            Else
                strPrintOption(cntRow, 0) = strTpePath & WsAtesaki.Cells(intTempRow(i), intOptCol(0))
            End If
            '�e���v���[�g���݃`�F�b�N��������
            '�w�肵�Ȃ��̏ꍇ�A�e���v���[�g�ƃe�v���T�C�Y���Ⴄ�ꍇ�A�G���[���b�Z�[�W
            If Dir(strPrintOption(cntRow, 0)) = "" Then
                If WsAtesaki.Cells(intTempRow(i), intOptCol(0)) = "�w�肵�Ȃ�" Then
                    MsgBox ERROR_MESSAGE_Default_Template & vbCrLf & vbCrLf & _
                           "�{�̂ɓ��ڂ���Ă���e�[�v �F " & strTapeWidth & " mm"
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
    '��������쐬�V�[�g�̃e���v���[�g���͋K�����X�V
    '================================================
    
    '�ϐ��錾
    Dim WsOption As Worksheet
    Set WsOption = Worksheets("option")
    Dim WsList As Worksheet
    Set WsList = Worksheets("List")
    '��������쐬
    Dim WsAtesaki As Worksheet
    Set WsAtesaki = Worksheets(strWsAtesaki)
    
    Dim strTempFile() As String         '�e���v���[�g�t�@�C�������i�[
    Dim strTempPath As String           '�e���v���[�g�̃p�X���i�[
    strTempPath = ThisWorkbook.Path & "\template"
    Dim strExt As String
    strExt = ".tpe"
    
    Dim intListRow(1) As Integer
    Dim intListCol(1) As Integer
    '���X�g�����͈�
    intListRow(0) = WsOption.Range("D11").Value
    intListCol(0) = WsOption.Range("D12").Value
    intListRow(1) = WsOption.Range("D13").Value
    intListCol(1) = WsOption.Range("D14").Value
    
    Dim i As Integer
    Dim cnt As Integer
    
    'List�V�[�g�̕����������N���A����
    '��ԏ�́u�w�肵�Ȃ��v�����炻��̓N���A���Ȃ�
    For i = intListRow(0) + 1 To intListRow(1)
        WsList.Cells(i, intListCol(0)).Clear
    Next i
    
    '�t�H���_�����񂳂����񂳂�
    '�e���v���[�g�t�H���_���́u.tpe�v�t�@�C�����擾����
    strTempFile = GetDirFile(strTempPath, strExt)
    
    'List�V�[�g�ɏ����o��
    '3�s�ڂ��珑���o��
    cnt = 3
    For i = LBound(strTempFile) To UBound(strTempFile)
        WsList.Cells(cnt, intListCol(0)).Value = strTempFile(i)
        cnt = cnt + 1
    Next i
    
    MsgBox "�e���v���[�g�t�@�C�����X�V���܂����B"
    
End Sub
