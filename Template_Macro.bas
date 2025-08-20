Attribute VB_Name = "Template_Macro"
Option Explicit


'���W���[�����x���萔
'------------------------------------------------------------------------
Public Const VERSION As String = "v1.0.0-beta"
Public Const MACRO_NAME As String = "Template_Macro"


'���W���[�����x���ϐ�
'------------------------------------------------------------------------
Private FO As New FileOjt
Private bm As New BookManager
Private dUtil As New DateUtility

Private gProjectFolderPath As String
Private gTempFolderPath As String
Private gReportFolderPath As String
Private gCurrentFolderPath As String


Public Sub MainContriller()

    '�ݒ�
    '------------------------------------------------------------------------
    Const PROC_NAME As String = "[MainContriller]"
    
    Application.ScreenUpdating = False
    
    gProjectFolderPath = FO.UpPath(ThisWorkbook.path)
    Call InitializeLogger(gProjectFolderPath)
    gTempFolderPath = gProjectFolderPath & "\output\temp"
    gReportFolderPath = gProjectFolderPath & "\output\reports"
    'gCurrentFolderPath = Environ("USERPROFILE") & "\Documents"
    gCurrentFolderPath = gProjectFolderPath & "\test" '�f�o�b�O�p�p�X
    
    logger.Info PROC_NAME & " macro name    : " & MACRO_NAME
    logger.Info PROC_NAME & " macro version : " & VERSION
    logger.DebugMsg PROC_NAME & " projectFolderPath : " & gProjectFolderPath
    logger.DebugMsg PROC_NAME & " tempFolderPath    : " & gTempFolderPath
    logger.DebugMsg PROC_NAME & " reportFolderPath  : " & gReportFolderPath
    
    logger.Info PROC_NAME & " �����̊J�n"
    
    ChDrive Left(gCurrentFolderPath, 1)
    ChDir gCurrentFolderPath
    
    On Error GoTo ErrHandler
        
    
    '�O����
    '------------------------------------------------------------------------
    Call waitMsgShow
    
    
    '�{����
    '------------------------------------------------------------------------
    Call TempMsgBox("�����𒆒f���܂����H")
    
    
    '�㏈��
    '------------------------------------------------------------------------
    Unload WaitMsg
    MsgBox "�������������܂����B", vbInformation, MACRO_NAME
    logger.Info PROC_NAME & " ����I��"
    Application.ScreenUpdating = True
    
    Exit Sub
    


ErrHandler:
    
    '�G���[����
    '------------------------------------------------------------------------
    logger.ErrorMsg PROC_NAME & " �G���[���� : " & Err.Description
    logger.WarnMsg PROC_NAME & " �����𒆒f���܂���"
    Unload WaitMsg
    MsgBox "�G���[���������܂����B" & vbCrLf & "�G���[���b�Z�[�W : " & Err.Description, vbExclamation, MACRO_NAME
    MsgBox "�����𒆒f���܂��B", vbExclamation, MACRO_NAME
    
End Sub

Private Sub TempMsgBox(ByVal message As String)
    
    Dim rc As Long
    
    rc = MsgBox(message, vbYesNo + vbQuestion, MACRO_NAME)
    If rc = vbYes Then
        MsgBox "�����𒆒f���܂��B", vbExclamation, MACRO_NAME
    Else
        MsgBox "�������������܂����B", vbInformation, MACRO_NAME
    End If

End Sub
