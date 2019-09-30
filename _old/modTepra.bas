'*****************************************************************************
' ���i�Ǘ����x����� �T���v���v���O���� for SPC10-API
' Copyright 2014 KING JIM CO.,LTD.
'*****************************************************************************

' ���W���[�����̎w��
Attribute VB_Name =  "modTepra"


' Windows API�̊֐����Ăяo��VBA�֐����`
Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" _
                    (ByVal lpModuleName As String) As Long
                       
Declare PtrSafe Function GetProcAddress Lib "kernel32" _
                    (ByVal hModule As Long, _
                     ByVal lpProcName As String) As Long
                         
Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As Long

Declare PtrSafe Function IsWow64Process Lib "kernel32" _
                    (ByVal hProcess As Long, _
                     ByRef Wow64Process As Long) As Long

' �f�[�^�J�n�s�̃I�t�Z�b�g
Public Const LINE_OFFSET As Integer = 19

' �Ώۃf�[�^�̍ő�s��
'Public Const MAX_LINE_COUNT As Integer = Range("D19").End(xlDown).Row
Public MAX_LINE_COUNT As Integer

' �I�v�V�����L��������
Public const Flg_Yes as String = "����"
Public const Flg_No as String = "���Ȃ�"

' �V�[�g��
Public Const strWsBihin As String = "���i�Ǘ��䒠"
Public Const strWsAtesaki As String = "��������쐬"

'�G���[���b�Z�[�W�̒�`
Public Const ERROR_MESSAGE_NO_PRINT_JOB = "������鍀�ڂ����͂���Ă��Ȃ���܂��͈���`�F�b�N�}�[�N�������Ă��Ȃ����ߤ���x��������ł��܂���"
Public Const ERROR_MESSAGE_GET_TAPE_WIDTH = "�e�[�v�����擾�ł��܂���B"
Public Const ERROR_MESSAGE_TPE_FILE_NOT_FOUND = "�e�[�v���ɍ��������C�A�E�g�����݂��܂���B"
Public Const ERROR_MESSAGE_RUN_PRINT = """SPC10.exe""���w�肵���ꏊ�ɑ��݂��܂���B�C���X�g�[������m�F���Ă��������B"

Public Const ERROR_MESSAGE_Job_Nothing = "����Ώۂ��I������Ă��܂���B"
Public Const ERROR_MESSAGE_Default_Template = "�y�e���v���[�g�w��z�Ły�w�肵�Ȃ��z��I�������ꍇ�́A24mm��36mm�̃e�[�v��{�̂ɓ���Ă��������B"
Public Const ERROR_MESSAGE_Template_Nothing = "�w�肳�ꂽ�e���v���[�g�����݂��܂���B"
Public Const ERROR_MESSAGE_Maisu_Nothing = "�s�ڂ́y�����z���w�肳��Ă��܂���B"
Public Const ERROR_MESSAGE_Muki_Nothing = "�s�ڂ́y�����z���w�肳��Ă��܂���B"
	
	
'==============================================================================
' OS��64�r�b�g�����ǂ����𔻕ʂ���֐��̒�`
'==============================================================================
Function IsWow64() As Boolean

'=============
'    Dim bIsWow64 As Long
'    Dim fnIsWow64Process As Long
'
'    ' ������
'    bIsWow64 = False
'    ' IsWow64Process�֐������݂��邩�ǂ������m�F����
'    fnIsWow64Process = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")
'
'    If (0 <> fnIsWow64Process) Then
'        ' IsWow64Process��Windows XP SP2���瓱�����ꂽ�֐��ł�
'        If 0 = IsWow64Process(GetCurrentProcess(), bIsWow64) Then
'            ' Windows XP �̌Â��o�[�W�����̏ꍇ�͂����ɗ��܂�
'        End If
'    End If
'=============

    Dim colItems As Object
    Dim itm As Object
    Dim ret As Boolean
     
    ret = False '������
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
' �e�[�v���擾�֐��̒�`
'==============================================================================
Function getTapeWidth(ByVal strFileName As String, ByRef strTapeType As String) As String
    
    Dim retValue As String
    
    ' �t�@�C����ǂݍ��ނ��߂̔z��
    Dim Arr()
    ReDim Preserve Arr(0)

    ' �I�u�W�F�N�g���쐬
    Dim obj As Object
    Set obj = CreateObject("ADODB.Stream")

    ' �I�u�W�F�N�g�ɕۑ�����f�[�^�̎�ނ𕶎���^�Ɏw�肷��
    obj.Type = adTypeText
    ' ������^�̃I�u�W�F�N�g�̕����R�[�h���w�肷��
    obj.Charset = "UTF-16"

    ' �I�u�W�F�N�g�̃C���X�^���X���쐬
    obj.Open

    ' �t�@�C������f�[�^��ǂݍ���
    obj.LoadFromFile (strFileName)

    ' �ŏI�s�܂Ń��[�v����
    Do While Not obj.EOS
        ' ���̍s��ǂݎ��
        Arr(UBound(Arr)) = obj.ReadText(adReadLine)
        ReDim Preserve Arr(UBound(Arr) + 1)
    Loop

    ' �I�u�W�F�N�g�����
    obj.Close

    ' ����������I�u�W�F�N�g���폜����
    Set obj = Nothing
    
    '-----------------------------------------------------------------------------
    ' �e�[�v���̎擾
    '-----------------------------------------------------------------------------
    ' �ǂݍ���1�s�ڂ̕�����𕪊��i��F0x04 18mm�j
    Dim strData As Variant
    strData = Split(Arr(0), " ")

    ' �e�[�v�����擾
    Dim strTapeWidth As String
    strTapeWidth = strData(0)

    ' �e�[�v���̐ݒ�
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
    ' �e�[�v��ނ̎擾
    '-----------------------------------------------------------------------------
    ' �ǂݍ���2�s�ڂ̕�����𕪊��i��F0x00 Standard tape�j
    Dim strTypeData As Variant
    strTypeData = Split(Arr(1), " ")

    ' �e�[�v��ނ��擾
    strTapeType = strTypeData(0)

    getTapeWidth = retValue
    
End Function

'==============================================================================
' �I�v�V���������񐶐��֐��̒�`
'==============================================================================
Function createPrintOption(pathTempl As String, pathCsv As String, printNum As Integer, blnHalfcut As Boolean, blnConfirmTapeWidth As Boolean, strPrintLog As String, strTapeWidth As String) As String

    Dim comStrg   As String
    Dim retValue  As Double
    Dim strOption As String
    
    ' TPE�t�@�C���̃t���p�X��,CSV�t�@�C���̃t���p�X��,�������
    strOption = pathTempl & "," & pathCsv & "," & printNum & ","
    
    ' �e�[�v���̃t�@�C���o��
    If (Len(strTapeWidth) > 0) Then
        strOption = strOption + "," + "/GT " + strTapeWidth
    End If
    
    ' �J�b�g�ݒ�
    If (blnHalfcut) Then
        strOption = strOption + "," + "/C -f -h"
    Else
        strOption = strOption + "," + "/C -f -hn"
    End If
    
    ' �e�[�v���m�F���b�Z�[�W��on/off�ݒ�
    If (blnConfirmTapeWidth) Then
        strOption = strOption + "," + "/TW -on"
    Else
        strOption = strOption + "," + "/TW -off"
    End If
    
    ' ������ʂ̃t�@�C���o��
    If (Len(strPrintLog) > 0) Then
        strOption = strOption + "," + "/L " + strPrintLog
    End If
   
    createPrintOption = strOption

End Function

'==============================================================================
' ������s�֐��̒�`
'==============================================================================
Function PrtSpc10Api(pathCmd As String, strOption As String, strPrinterName As String) As Double
    
    Dim comStrg  As String
    Dim retValue As Double
    
    ' ����R�}���h
    If (Len(strPrinterName) > 0) Then
        ' /pt�I�v�V����
        comStrg = pathCmd & " " & "/pt " & Chr(34) & strOption & Chr(34) & " " & Chr(34) & strPrinterName & Chr(34)
    Else
        ' /p�I�v�V����
        comStrg = pathCmd & " " & "/p " & Chr(34) & strOption & Chr(34)
    End If
    
    On Error Resume Next
    
    retValue = Shell(comStrg, vbHide)
    
    PrtSpc10Api = retValue
    
End Function


Private sub ErrorConst()

	
end sub

