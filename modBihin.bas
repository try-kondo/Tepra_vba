'*****************************************************************************
' ���i�Ǘ����x����� �T���v���v���O���� for SPC10-API
' Copyright 2014 KING JIM CO.,LTD.
'*****************************************************************************

' ���W���[�����̎w��
Attribute VB_Name =  "modBihin"

'==============================================================================
' CSV�t�@�C�����쐬���A������s�֐����Ăяo��
'==============================================================================
Public Sub Bihin_main()

    Dim strExePathName As String
    Dim strTextPathName As String
    Dim strCsvPathName As String
    Dim strPrintLogPathName As String
    Dim strTpePathName As String
    Dim strOption As String
    Dim dblRetValue As Double
    
    '�Ώۂ̃W���u
    Dim strPrintjob() As String
    
    Dim strColName As String
    Dim strColDel As String
    
    Dim WsOption As Worksheet
    Set WsOption = Worksheets("option")
    
    Dim WsList As Worksheet
    Set WsList = Worksheets("List")
    
    '���i�Ǘ��䒠
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
'    strTpePathName = ThisWorkbook.Path & "\bihin_18.tpe"              ' TPE�t�@�C�����w�肷��i�f�t�H���g:18mm�j
    ' TPE�t�@�C�����w�肷��i�f�t�H���g:12mm - 1Line�j
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
    ' ����Ώۂ̊m�F
    '-----------------------------------------------------------------------------
    strPrintjob() = getPrintJobCount()
    
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

    ' TPE�t�@�C�����w��
    '�J��������L���`�F�b�N
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
        ' TPE�t�@�C�������݂��Ȃ��ꍇ
        MsgBox ERROR_MESSAGE_TPE_FILE_NOT_FOUND
        Exit Sub
    End If
    
    '-----------------------------------------------------------------------------
    ' CSV�t�@�C���̍쐬
    '-----------------------------------------------------------------------------

    On Error Resume Next
    
    ' �����l�ݒ�
    fileNo = FreeFile
    
'    ' ��Ж�
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
        'CSV�t�@�C���o��
        Print #fileNo, csvStrAll
    Next i
    
    Close #fileNo

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

End Sub

'==============================================================================
' ����W���u�����擾����
'==============================================================================
Private Function getPrintJobCount() As String()

    Dim i As Integer
    
    Dim strPrintjob() As String     'CSV�����o���p
    Dim intPrintCol() As Integer    '����Ώۂ̃J�������i�[
    Dim intMaxCol As Integer        '�\�̍ő��
    intMaxCol = Range("D19").End(xlToRight).Column
    
    Dim intPrintRow() As Integer
    
    Dim cnt As Integer
    
    Dim j As Integer
    
    '���i�Ǘ��䒠
    Dim WsBihin As Worksheet
    Set WsBihin = Worksheets(strWsBihin)
    
    '�Ώۂ̃J�����̌���
    cnt = 0
    ReDim intPrintCol(intMaxCol)
    For i = 4 To intMaxCol
        If WsBihin.Cells(LINE_OFFSET - 1, i).Value = "��" Then
            intPrintCol(cnt) = i
            cnt = cnt + 1
        End If
    Next i
    
    '�J�����������I������Ă��Ȃ��Ƃ�
    If cnt = 0 Then
        MsgBox "����Ώۂ̗񂪑I������Ă��܂���"
        End
    End If
    
    ReDim Preserve intPrintCol(cnt - 1)
    '�Ώۍs�̌���
    cnt = 0
    ReDim intPrintRow(MAX_LINE_COUNT)
    For i = 20 To MAX_LINE_COUNT
        If WsBihin.Cells(i, 3).Value = "��" Then
            intPrintRow(cnt) = i
            cnt = cnt + 1
        End If
    Next i
    
    '�s�������I������Ă��Ȃ��Ƃ�
    If cnt = 0 Then
        MsgBox "����Ώۂ̍s���I������Ă��܂���"
        End
    End If
    
    ReDim Preserve intPrintRow(cnt - 1)
    
    'CSV�쐬�p
    ReDim strPrintjob(UBound(intPrintRow), UBound(intPrintCol) * 2 + 1) '�J���������������ނ���񂾂��Q�{����
    For i = 0 To UBound(intPrintRow)
        cnt = 0
        For j = 0 To UBound(strPrintjob, 2) Step 2
            '�J��������������
            strPrintjob(i, j) = WsBihin.Cells(LINE_OFFSET, intPrintCol(cnt)).Value
            '���g��������
            strPrintjob(i, j + 1) = WsBihin.Cells(intPrintRow(i), intPrintCol(cnt)).Value
            
            cnt = cnt + 1
        Next j
        
    Next i
    
    getPrintJobCount = strPrintjob
    
End Function

